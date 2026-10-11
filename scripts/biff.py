"""Minimal BIFF8 (.xls) reader/writer helpers for patching cell values in place.

The workbook stream is parsed into records; edits replace individual cell records, append strings
to the shared string table and flip form-control check boxes. Offsets that depend on record sizes
(BOUNDSHEET, INDEX, DBCELL, EXTSST) are recomputed when the stream is serialised again.
"""
import struct

BOF, EOF_, BOUNDSHEET, SST, EXTSST, CONTINUE, INDEX, DBCELL, ROW = 0x0809, 0x000A, 0x0085, 0x00FC, 0x00FF, 0x003C, 0x020B, 0x00D7, 0x0208
BLANK, MULBLANK, LABELSST, NUMBER, RK, MULRK, BOOLERR, FORMULA, LABEL, RSTRING, STRING = 0x0201, 0x00BE, 0x00FD, 0x0203, 0x027E, 0x00BD, 0x0205, 0x0006, 0x0204, 0x00D6, 0x0207
OBJ, MSODRAWING, TXO = 0x005D, 0x00EC, 0x01B6
CELL_SIDS = {BLANK, MULBLANK, LABELSST, NUMBER, RK, MULRK, BOOLERR, FORMULA, LABEL, RSTRING}
MAX_REC = 8224


class Rec:
    __slots__ = ("sid", "data")

    def __init__(self, sid, data):
        self.sid, self.data = sid, bytes(data)

    def __repr__(self):
        return f"Rec(0x{self.sid:04X}, {len(self.data)})"


def parse(stream):
    recs, pos = [], 0
    while pos + 4 <= len(stream):
        sid, ln = struct.unpack_from("<HH", stream, pos)
        recs.append(Rec(sid, stream[pos + 4:pos + 4 + ln]))
        pos += 4 + ln
    return recs


def serialise(recs):
    out = bytearray()
    for r in recs:
        out += struct.pack("<HH", r.sid, len(r.data)) + r.data
    return bytes(out)


def positions(recs):
    pos, out = 0, []
    for r in recs:
        out.append(pos)
        pos += 4 + len(r.data)
    return out


def substreams(recs):
    """Return [(start_index, end_index)] for each BOF..EOF substream (globals first)."""
    out, start, depth = [], None, 0
    for i, r in enumerate(recs):
        if r.sid == BOF:
            if depth == 0:
                start = i
            depth += 1
        elif r.sid == EOF_:
            depth -= 1
            if depth == 0:
                out.append((start, i))
    return out


def cell_rc(rec):
    """(row, first_col, last_col) for a cell record."""
    row, col = struct.unpack_from("<HH", rec.data, 0)
    if rec.sid in (MULBLANK, MULRK):
        last = struct.unpack_from("<H", rec.data, len(rec.data) - 2)[0]
        return row, col, last
    return row, col, col


def read_unicode(buf, pos, cch=None, len_bytes=2):
    """XLUnicodeString without continuation handling (for short strings inside one record)."""
    if cch is None:
        cch = struct.unpack_from("<H" if len_bytes == 2 else "<B", buf, pos)[0]
        pos += len_bytes
    flags = buf[pos]
    pos += 1
    if flags & 1:
        s = buf[pos:pos + 2 * cch].decode("utf-16-le")
        pos += 2 * cch
    else:
        s = buf[pos:pos + cch].decode("latin-1")
        pos += cch
    return s, pos


def sst_strings(recs, sst_index):
    """Decode all strings of the SST (handles CONTINUE splits). Returns (strings, record_count)."""
    chunks = [recs[sst_index].data]
    j = sst_index + 1
    while j < len(recs) and recs[j].sid == CONTINUE:
        chunks.append(recs[j].data)
        j += 1
    total, unique = struct.unpack_from("<II", chunks[0], 0)
    ci, pos = 0, 8
    out = []

    def need(n):
        nonlocal ci, pos
        if pos >= len(chunks[ci]):
            ci += 1
            pos = 0

    for _ in range(unique):
        need(1)
        buf = chunks[ci]
        cch = struct.unpack_from("<H", buf, pos)[0]
        flags = buf[pos + 2]
        pos += 3
        crun = cbext = 0
        if flags & 8:
            crun = struct.unpack_from("<H", buf, pos)[0]
            pos += 2
        if flags & 4:
            cbext = struct.unpack_from("<I", buf, pos)[0]
            pos += 4
        high = flags & 1
        chars = []
        remaining = cch
        while remaining:
            buf = chunks[ci]
            if pos >= len(buf):
                ci += 1
                pos = 0
                buf = chunks[ci]
                high = buf[0] & 1
                pos = 1
            width = 2 if high else 1
            avail = (len(buf) - pos) // width
            take = min(avail, remaining)
            raw = buf[pos:pos + take * width]
            chars.append(raw.decode("utf-16-le") if high else raw.decode("latin-1"))
            pos += take * width
            remaining -= take
        skip = 4 * crun + cbext
        while skip:
            buf = chunks[ci]
            if pos >= len(buf):
                ci += 1
                pos = 0
                continue
            take = min(skip, len(buf) - pos)
            pos += take
            skip -= take
        out.append("".join(chars))
    return out, j - sst_index, total


# ---------------------------------------------------------------- drawing objects and form controls
def officeart_shapes(buf):
    """Walk OfficeArt records; return a list of shapes in order with spid, name and anchor."""
    shapes, cur = [], None
    pos = 0

    def walk(start, end):
        nonlocal cur
        p = start
        while p + 8 <= end:
            verinst, rtype, rlen = struct.unpack_from("<HHI", buf, p)
            ver = verinst & 0x0F
            body = p + 8
            if ver == 0x0F:  # container
                if rtype == 0xF004:  # SpContainer
                    cur = {"spid": None, "name": None, "anchor": None}
                    shapes.append(cur)
                walk(body, min(body + rlen, end))
                p = body + rlen
                continue
            if rtype == 0xF00A and cur is not None:  # FSP
                cur["spid"] = struct.unpack_from("<I", buf, body)[0]
            elif rtype == 0xF00B and cur is not None:  # FOPT
                nprops = verinst >> 4
                cplx = body + 6 * nprops
                for k in range(nprops):
                    opid, op = struct.unpack_from("<HI", buf, body + 6 * k)
                    pid, is_cplx = opid & 0x3FFF, opid & 0x8000
                    if is_cplx:
                        if pid == 0x0380:
                            cur["name"] = buf[cplx:cplx + op].decode("utf-16-le").rstrip("\x00")
                        cplx += op
            elif rtype == 0xF010 and cur is not None and rlen >= 18:  # ClientAnchor
                _, c1, dx1, r1, dy1, c2, dx2, r2, dy2 = struct.unpack_from("<HHHHHHHHH", buf, body)
                cur["anchor"] = (r1, c1, r2, c2)
            p = body + rlen

    walk(0, len(buf))
    return shapes


def obj_info(rec):
    """Decode the parts of an OBJ record needed for check boxes."""
    d, p, info = rec.data, 0, {"subs": []}
    while p + 4 <= len(d):
        ft, cb = struct.unpack_from("<HH", d, p)
        if ft == 0x0015:
            info["id"], info["ot"] = struct.unpack_from("<HH", d, p + 4)[::-1][::-1][0:2][::-1][::-1] if False else (struct.unpack_from("<H", d, p + 6)[0], struct.unpack_from("<H", d, p + 4)[0])
        elif ft == 0x0014:  # FtCblsFmla: linked cell
            info["link_raw"] = d[p + 4:p + 4 + cb]
        elif ft == 0x0012:  # FtCblsData
            info["checked_off"] = p + 4
            info["checked"] = struct.unpack_from("<H", d, p + 4)[0]
        elif ft == 0x0004:  # FtMacro
            info["macro_raw"] = d[p + 4:p + 4 + cb]
        info["subs"].append((ft, cb))
        if ft == 0x0000:
            break
        p += 4 + cb
    return info


def decode_ref(raw):
    """Return (row, col) from the first ptgRef/ptgRef3d found in an ObjFmla blob."""
    for k in range(len(raw)):
        t = raw[k]
        if t in (0x24, 0x44, 0x64) and k + 5 <= len(raw):
            rw, col = struct.unpack_from("<HH", raw, k + 1)
            return rw, col & 0x3FFF
        if t in (0x3A, 0x5A, 0x7A) and k + 7 <= len(raw):
            ixti, rw, col = struct.unpack_from("<HHH", raw, k + 1)
            return rw, col & 0x3FFF
    return None
