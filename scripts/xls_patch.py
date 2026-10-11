"""Patch cell values and check boxes in a macro-enabled BIFF8 .xls file without re-saving it in Excel.

Only the records that hold the edited cells change. The shared string table gets new strings
appended, and the size-dependent offsets (BOUNDSHEET, INDEX, DBCELL, EXTSST) are recomputed. The
compound file is patched in place: the Workbook stream's sector chain is extended and every other
stream (including the VBA project) keeps its original bytes.
"""
import datetime
import re
import struct

import olefile

import biff
from biff import (BLANK, BOOLERR, BOUNDSHEET, CELL_SIDS, CONTINUE, DBCELL, EXTSST, FORMULA, INDEX, LABELSST, MAX_REC,
                  MULBLANK, MULRK, NUMBER, OBJ, ROW, SST, STRING, Rec)

ENDOFCHAIN, FREESECT, FATSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD


def a1(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, col - 1


def excel_date(d):
    return float((d - datetime.date(1899, 12, 30)).days)


def xl_unicode_rich(s):
    """XLUnicodeRichExtendedString without formatting runs (for the SST)."""
    try:
        body = s.encode("latin-1")
        return struct.pack("<HB", len(s), 0) + body
    except UnicodeEncodeError:
        return struct.pack("<HB", len(s), 1) + s.encode("utf-16-le")


class XlsPatcher:
    def __init__(self, path):
        self.path = path
        ole = olefile.OleFileIO(path)
        self.stream = ole.openstream("Workbook").read()
        ole.close()
        self.recs = biff.parse(self.stream)
        assert biff.serialise(self.recs) == self.stream
        subs = biff.substreams(self.recs)
        self.names = []
        for r in self.recs[subs[0][0]:subs[0][1]]:
            if r.sid == BOUNDSHEET:
                self.names.append(biff.read_unicode(r.data, 6, len_bytes=1)[0])
        self.sst_i = next(i for i, r in enumerate(self.recs) if r.sid == SST)
        self.strings, self.sst_nrec, self.sst_total = biff.sst_strings(self.recs, self.sst_i)
        self.lookup = {}
        for i, s in enumerate(self.strings):
            self.lookup.setdefault(s, i)
        self.n_orig_strings = len(self.strings)
        self.added_refs = 0
        self.log = []

    # ------------------------------------------------------------------ locating records
    def _sheet_range(self, sheet):
        k = self.names.index(sheet) + 1
        return biff.substreams(self.recs)[k]

    def _find_cell(self, sheet, row, col):
        s, e = self._sheet_range(sheet)
        for i in range(s, e):
            r = self.recs[i]
            if r.sid in CELL_SIDS:
                rr, c1, c2 = biff.cell_rc(r)
                if rr == row and c1 <= col <= c2:
                    return i
        raise KeyError(f"{sheet}!{row},{col}: no cell record (cell has never been formatted)")

    def _xf_of(self, rec, col):
        if rec.sid in (MULBLANK, MULRK):
            row, c1 = struct.unpack_from("<HH", rec.data, 0)
            if rec.sid == MULBLANK:
                return struct.unpack_from("<H", rec.data, 4 + 2 * (col - c1))[0]
            return struct.unpack_from("<H", rec.data, 4 + 6 * (col - c1))[0]
        return struct.unpack_from("<H", rec.data, 4)[0]

    def _replace(self, sheet, ref, make):
        """Replace the cell at ref with the record built by make(row, col, xf)."""
        row, col = a1(ref)
        i = self._find_cell(sheet, row, col)
        old = self.recs[i]
        xf = self._xf_of(old, col)
        new = make(row, col, xf)
        if old.sid == MULBLANK:
            _, c1 = struct.unpack_from("<HH", old.data, 0)
            xfs = list(struct.unpack_from(f"<{(len(old.data) - 6) // 2}H", old.data, 4))
            c2 = c1 + len(xfs) - 1
            parts = []
            if col > c1:
                parts.append(self._blank_run(row, c1, xfs[:col - c1]))
            parts.append(new)
            if col < c2:
                parts.append(self._blank_run(row, col + 1, xfs[col - c1 + 1:]))
            self.recs[i:i + 1] = parts
        elif old.sid == MULRK:
            raise NotImplementedError("MULRK split not needed for this form")
        else:
            if old.sid == LABELSST:
                self.added_refs -= 1
            drop = 1
            if old.sid == FORMULA and i + 1 < len(self.recs) and self.recs[i + 1].sid == STRING:
                drop = 2
            self.recs[i:i + drop] = [new]
        self.log.append((sheet, ref, old.sid))

    @staticmethod
    def _blank_run(row, c1, xfs):
        if len(xfs) == 1:
            return Rec(BLANK, struct.pack("<HHH", row, c1, xfs[0]))
        return Rec(MULBLANK, struct.pack("<HH", row, c1) + struct.pack(f"<{len(xfs)}H", *xfs) + struct.pack("<H", c1 + len(xfs) - 1))

    # ------------------------------------------------------------------ public edits
    def set_text(self, sheet, ref, text):
        idx = self.lookup.get(text)
        if idx is None:
            idx = len(self.strings)
            self.strings.append(text)
            self.lookup[text] = idx
        self.added_refs += 1
        self._replace(sheet, ref, lambda r, c, xf: Rec(LABELSST, struct.pack("<HHHI", r, c, xf, idx)))

    def set_number(self, sheet, ref, value):
        self._replace(sheet, ref, lambda r, c, xf: Rec(NUMBER, struct.pack("<HHHd", r, c, xf, float(value))))

    def set_date(self, sheet, ref, d):
        self.set_number(sheet, ref, excel_date(d))

    def set_bool(self, sheet, ref, value):
        self._replace(sheet, ref, lambda r, c, xf: Rec(BOOLERR, struct.pack("<HHHBB", r, c, xf, 1 if value else 0, 0)))

    def set_formula_result(self, sheet, ref, value):
        """Update the cached result of an existing formula (number, bool or '' for an empty string)."""
        row, col = a1(ref)
        i = self._find_cell(sheet, row, col)
        rec = self.recs[i]
        assert rec.sid == FORMULA, (sheet, ref, hex(rec.sid))
        d = bytearray(rec.data)
        if isinstance(value, bool):
            d[6:14] = bytes([1, 0, 1 if value else 0, 0, 0, 0, 0xFF, 0xFF])
        elif value == "":
            d[6:14] = bytes([3, 0, 0, 0, 0, 0, 0xFF, 0xFF])
        else:
            d[6:14] = struct.pack("<d", float(value))
        drop = 2 if i + 1 < len(self.recs) and self.recs[i + 1].sid == STRING else 1
        self.recs[i:i + drop] = [Rec(FORMULA, d)]

    def checkbox_ids(self, sheet):
        s, e = self._sheet_range(sheet)
        out = {}
        for i in range(s, e):
            if self.recs[i].sid == OBJ:
                info = biff.obj_info(self.recs[i])
                if info.get("ot") == 11:
                    link = biff.decode_ref(info["link_raw"]) if "link_raw" in info else None
                    out[info["id"]] = (i, link)
        return out

    def set_checkbox(self, sheet, obj_id, checked=True):
        i, _ = self.checkbox_ids(sheet)[obj_id]
        info = biff.obj_info(self.recs[i])
        d = bytearray(self.recs[i].data)
        struct.pack_into("<H", d, info["checked_off"], 1 if checked else 0)
        self.recs[i] = Rec(OBJ, d)
        self.log.append((sheet, f"checkbox {obj_id}", OBJ))

    def checkbox_for_link(self, sheet, ref):
        row, col = a1(ref)
        for oid, (_, link) in self.checkbox_ids(sheet).items():
            if link == (row, col):
                return oid
        raise KeyError(f"no check box linked to {sheet}!{ref}")

    # ------------------------------------------------------------------ rebuilding the stream
    def _rebuild_sst(self):
        old = self.recs[self.sst_i:self.sst_i + self.sst_nrec]
        chunks = [bytearray(r.data) for r in old]
        new_positions = []  # (chunk index, offset within chunk data) of each appended string
        for s in self.strings[self.n_orig_strings:]:
            blob = xl_unicode_rich(s)
            if len(chunks[-1]) + len(blob) > MAX_REC:
                chunks.append(bytearray())
            new_positions.append((len(chunks) - 1, len(chunks[-1])))
            chunks[-1] += blob
        total = self.sst_total + self.added_refs
        struct.pack_into("<II", chunks[0], 0, total, len(self.strings))
        recs = [Rec(SST, chunks[0])] + [Rec(CONTINUE, c) for c in chunks[1:]]
        self.recs[self.sst_i:self.sst_i + self.sst_nrec] = recs
        self.sst_nrec = len(recs)
        return new_positions

    def to_stream(self):
        new_positions = self._rebuild_sst()
        ext_i = next(i for i, r in enumerate(self.recs) if r.sid == EXTSST)
        dsst = struct.unpack_from("<H", self.recs[ext_i].data, 0)[0]
        nbuckets = (len(self.strings) + dsst - 1) // dsst
        old_ext = self.recs[ext_i].data
        old_buckets = (len(old_ext) - 2) // 8
        self.recs[ext_i] = Rec(EXTSST, old_ext[:2] + old_ext[2:2 + 8 * old_buckets] + bytes(8 * (nbuckets - old_buckets)))
        # positions are final once sizes are fixed (offset fields below never change a record's size)
        pos = biff.positions(self.recs)
        ext = bytearray(self.recs[ext_i].data)
        for b in range(old_buckets, nbuckets):
            k = b * dsst - self.n_orig_strings
            assert k >= 0, "new bucket must start at an appended string"
            ci, off = new_positions[k]
            rec_pos = pos[self.sst_i + ci]
            struct.pack_into("<IHH", ext, 2 + 8 * b, rec_pos + 4 + off, 4 + off, 0)
        self.recs[ext_i] = Rec(EXTSST, ext)
        subs = biff.substreams(self.recs)
        # BOUNDSHEET -> BOF positions
        k = 0
        for i in range(subs[0][0], subs[0][1]):
            if self.recs[i].sid == BOUNDSHEET:
                k += 1
                d = bytearray(self.recs[i].data)
                struct.pack_into("<I", d, 0, pos[subs[k][0]])
                self.recs[i] = Rec(BOUNDSHEET, d)
        # INDEX and DBCELL per sheet
        for s, e in subs[1:]:
            idx = [i for i in range(s, e) if self.recs[i].sid == INDEX]
            if not idx:
                continue
            dbs = [i for i in range(s, e) if self.recs[i].sid == DBCELL]
            prev = idx[0]
            for di in dbs:
                rows = [i for i in range(prev + 1, di) if self.recs[i].sid == ROW]
                cells = [i for i in range(prev + 1, di) if self.recs[i].sid in CELL_SIDS]
                firsts = {}
                for ci in cells:
                    firsts.setdefault(struct.unpack_from("<H", self.recs[ci].data, 0)[0], ci)
                rownums = [struct.unpack_from("<H", self.recs[i].data, 0)[0] for i in rows]
                rg, prevp = [], (pos[rows[1]] if len(rows) > 1 else pos[rows[0]] + 4 + len(self.recs[rows[0]].data))
                for rn in rownums:
                    fc = firsts.get(rn)
                    if fc is None:
                        rg.append(0)
                        continue
                    rg.append(pos[fc] - prevp)
                    prevp = pos[fc]
                d = struct.pack("<I", pos[di] - pos[rows[0]]) + struct.pack(f"<{len(rg)}H", *rg)
                assert len(d) == len(self.recs[di].data), "row block changed shape"
                self.recs[di] = Rec(DBCELL, d)
                prev = di
            d = bytearray(self.recs[idx[0]].data)
            dcw = [i for i in range(s, e) if self.recs[i].sid == 0x0055]
            if dcw:
                struct.pack_into("<I", d, 12, pos[dcw[0]])
            for j, di in enumerate(dbs):
                struct.pack_into("<I", d, 16 + 4 * j, pos[di])
            self.recs[idx[0]] = Rec(INDEX, d)
        return biff.serialise(self.recs)

    def save(self, out_path):
        write_stream(self.path, out_path, "Workbook", self.to_stream())


# ---------------------------------------------------------------------- compound file
def write_stream(src, dst, name, data):
    """Copy src to dst, replacing one large stream; all other bytes are kept."""
    buf = bytearray(open(src, "rb").read())
    hdr = buf[:512]
    assert hdr[:8] == bytes.fromhex("D0CF11E0A1B11AE1")
    ssz = 1 << struct.unpack_from("<H", hdr, 0x1E)[0]
    assert ssz == 512, "only version-3 files are handled"
    nfat = struct.unpack_from("<I", hdr, 0x2C)[0]
    first_dir = struct.unpack_from("<I", hdr, 0x30)[0]
    cutoff = struct.unpack_from("<I", hdr, 0x38)[0]
    ndifat = struct.unpack_from("<I", hdr, 0x48)[0]
    assert ndifat == 0 and nfat <= 109, "DIFAT sectors are not handled"
    fat_secs = list(struct.unpack_from(f"<{nfat}I", hdr, 0x4C))
    off = lambda sec: 512 + sec * 512  # noqa: E731
    fat = []
    for fs in fat_secs:
        fat += list(struct.unpack_from("<128I", buf, off(fs)))

    def chain(start):
        out = []
        while start not in (ENDOFCHAIN, FREESECT):
            out.append(start)
            start = fat[start]
        return out

    # find the directory entry
    dir_secs = chain(first_dir)
    entry_off = None
    for ds in dir_secs:
        for k in range(4):
            eo = off(ds) + 128 * k
            nlen = struct.unpack_from("<H", buf, eo + 64)[0]
            ename = buf[eo:eo + max(nlen - 2, 0)].decode("utf-16-le")
            if ename == name and buf[eo + 66] == 2:
                entry_off = eo
    assert entry_off is not None, f"stream {name} not found"
    start = struct.unpack_from("<I", buf, entry_off + 116)[0]
    old_size = struct.unpack_from("<I", buf, entry_off + 120)[0]
    assert old_size >= cutoff and len(data) >= cutoff, "only regular (non-mini) streams are handled"
    old_chain = chain(start)
    need = (len(data) + 511) // 512
    new_chain = old_chain[:need]
    for sec in old_chain[need:]:
        fat[sec] = FREESECT
    nsec_file = (len(buf) - 512 + 511) // 512
    while len(new_chain) < need:
        free = next((s for s in range(nsec_file) if fat[s] == FREESECT and s not in new_chain), None)
        if free is None:
            free = nsec_file
            nsec_file += 1
        while free >= len(fat) or nsec_file > len(fat):  # grow the FAT by one sector
            fs = nsec_file
            nsec_file += 1
            fat += [FREESECT] * 128
            fat[fs] = FATSECT
            fat_secs.append(fs)
            assert len(fat_secs) <= 109
        new_chain.append(free)
    for a, b in zip(new_chain, new_chain[1:]):
        fat[a] = b
    fat[new_chain[-1]] = ENDOFCHAIN
    # grow the file and write the stream
    if len(buf) < 512 + nsec_file * 512:
        buf += bytes(512 + nsec_file * 512 - len(buf))
    padded = data + bytes(need * 512 - len(data))
    for j, sec in enumerate(new_chain):
        buf[off(sec):off(sec) + 512] = padded[j * 512:(j + 1) * 512]
    # write FAT sectors and header
    for j, fs in enumerate(fat_secs):
        buf[off(fs):off(fs) + 512] = struct.pack("<128I", *fat[j * 128:(j + 1) * 128])
    struct.pack_into("<I", buf, 0x2C, len(fat_secs))
    difat = fat_secs + [FREESECT] * (109 - len(fat_secs))
    struct.pack_into("<109I", buf, 0x4C, *difat)
    struct.pack_into("<I", buf, entry_off + 116, new_chain[0])
    struct.pack_into("<II", buf, entry_off + 120, len(data), 0)
    open(dst, "wb").write(bytes(buf))
