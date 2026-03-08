"""make_rawcopy.py — HWPX raw-copy ZIP builder for flag_bits investigation.

Creates test_d_rawcopy.hwpx where:
  - All entries EXCEPT Contents/section0.xml: raw compressed bytes copied
    directly from the original ZIP, preserving flag_bits, CRC, compress_size.
  - Contents/section0.xml: modified XML from form_pass1.hwpx (re-compressed).

The script also prints a flag_bits comparison table between:
  1. Original form_to_fillout.hwpx
  2. Current form_pass1.hwpx  (produced by hwpx_editor.py -> writestr)
  3. test_d_rawcopy.hwpx       (this script's output)

Usage:
    python3 tools/make_rawcopy.py
"""

import io
import os
import struct
import zipfile
import zlib

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
REPO_ROOT    = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd"
ORIGINAL_HWPX = os.path.join(REPO_ROOT, "form_to_fillout.hwpx")
PASS1_HWPX    = os.path.join(REPO_ROOT, "hwpx-generator/output/filled/form_pass1.hwpx")
OUTPUT_DIR    = os.path.join(REPO_ROOT, "hwpx-generator/output/debug")
OUTPUT_HWPX   = os.path.join(OUTPUT_DIR, "test_d_rawcopy.hwpx")

SECTION_ENTRY = "Contents/section0.xml"


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _get_data_offset(orig_path, header_offset):
    """Return the byte offset where compressed data begins for a local entry."""
    with open(orig_path, "rb") as f:
        f.seek(header_offset)
        sig = f.read(4)
        if sig != b"PK\x03\x04":
            raise ValueError(f"Bad local file header sig at offset {header_offset:#x}")
        f.seek(header_offset + 26)
        fname_len, extra_len = struct.unpack("<HH", f.read(4))
    return header_offset + 30 + fname_len + extra_len


def read_raw_compressed(orig_path, info):
    """Read the raw (already-compressed) bytes for a ZipInfo entry.

    Bypasses Python's decompressor so the exact compressed stream,
    flag_bits, CRC, and compress_size are preserved.
    """
    data_offset = _get_data_offset(orig_path, info.header_offset)
    with open(orig_path, "rb") as f:
        f.seek(data_offset)
        return f.read(info.compress_size)


def _dos_time(date_time):
    """Convert Python ZipInfo date_time tuple to MS-DOS mod_time and mod_date."""
    yr, mo, day, hr, mn, sc = date_time
    mod_time = (hr << 11) | (mn << 5) | (sc >> 1)
    mod_date = ((yr - 1980) << 9) | (mo << 5) | day
    return mod_time, mod_date


# ---------------------------------------------------------------------------
# Minimal raw ZIP writer
# ---------------------------------------------------------------------------
# ZIP record formats (all little-endian):
#
# Local file header (30 bytes fixed + fname + extra):
#   PK\x03\x04  version_needed(H)  flags(H)  method(H)
#   mod_time(H)  mod_date(H)  crc32(I)  compress_size(I)  file_size(I)
#   fname_len(H)  extra_len(H)
#
# Central directory header (46 bytes fixed + fname + extra + comment):
#   PK\x01\x02  version_made(H)  version_needed(H)  flags(H)  method(H)
#   mod_time(H)  mod_date(H)  crc32(I)  compress_size(I)  file_size(I)
#   fname_len(H)  extra_len(H)  comment_len(H)  disk_start(H)
#   int_attr(H)  ext_attr(I)  local_header_offset(I)
#
# End of central directory (22 bytes fixed + comment):
#   PK\x05\x06  disk_num(H)  disk_cd_start(H)
#   entries_this_disk(H)  total_entries(H)
#   cd_size(I)  cd_offset(I)  comment_len(H)

_LFH_FMT  = struct.Struct("<HHHHHIIIHH")   # 10 fields, 26 bytes
_CDH_FMT  = struct.Struct("<HHHHHHIIIHHHHHII")  # 16 fields, 42 bytes
_EOCD_FMT = struct.Struct("<HHHHIIH")       # 7 fields, 18 bytes


class RawZipWriter:
    """Minimal ZIP writer that can inject already-compressed streams verbatim.

    add_raw(orig_path, info)  -- preserves flag_bits / CRC / compress_size
    add_data(info, data, ...)  -- normal compression (flag_bits becomes 0)
    close()                    -- finalize and flush to disk
    """

    def __init__(self, path):
        self._path = path
        self._buf  = io.BytesIO()
        self._cd   = []   # central directory entries

    # ------------------------------------------------------------------
    # Internal: write one local file header + data
    # ------------------------------------------------------------------
    def _write_entry(self, fname, date_time, compress_type,
                     flag_bits, crc, compress_size, file_size, raw_data):
        fname_bytes = fname.encode("utf-8")
        extra       = b""
        offset      = self._buf.tell()
        mod_time, mod_date = _dos_time(date_time)

        self._buf.write(b"PK\x03\x04")
        self._buf.write(_LFH_FMT.pack(
            20,             # version needed
            flag_bits,
            compress_type,
            mod_time,
            mod_date,
            crc,
            compress_size,
            file_size,
            len(fname_bytes),
            len(extra),
        ))
        self._buf.write(fname_bytes)
        self._buf.write(extra)
        self._buf.write(raw_data)

        self._cd.append(dict(
            fname_bytes  = fname_bytes,
            date_time    = date_time,
            compress_type= compress_type,
            flag_bits    = flag_bits,
            crc          = crc,
            compress_size= compress_size,
            file_size    = file_size,
            local_offset = offset,
        ))

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def add_raw(self, orig_path, info):
        """Copy an entry verbatim from orig_path (preserves flag_bits etc.)."""
        raw_data = read_raw_compressed(orig_path, info)
        self._write_entry(
            fname         = info.filename,
            date_time     = info.date_time,
            compress_type = info.compress_type,
            flag_bits     = info.flag_bits,
            crc           = info.CRC,
            compress_size = info.compress_size,
            file_size     = info.file_size,
            raw_data      = raw_data,
        )

    def add_data(self, template_info, data, compress_type=zipfile.ZIP_DEFLATED):
        """Add new/modified uncompressed data (flag_bits will be 0)."""
        if compress_type == zipfile.ZIP_DEFLATED:
            cobj     = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15)
            raw_data = cobj.compress(data) + cobj.flush()
        else:
            raw_data = data

        crc32 = zlib.crc32(data) & 0xFFFFFFFF
        self._write_entry(
            fname         = template_info.filename,
            date_time     = template_info.date_time,
            compress_type = compress_type,
            flag_bits     = 0,
            crc           = crc32,
            compress_size = len(raw_data),
            file_size     = len(data),
            raw_data      = raw_data,
        )

    def close(self):
        """Write the central directory and EOCD, then flush to disk."""
        cd_offset = self._buf.tell()

        for e in self._cd:
            fname_bytes = e["fname_bytes"]
            extra       = b""
            comment     = b""
            mod_time, mod_date = _dos_time(e["date_time"])

            self._buf.write(b"PK\x01\x02")
            self._buf.write(_CDH_FMT.pack(
                20,                  # version made by
                20,                  # version needed
                e["flag_bits"],
                e["compress_type"],
                mod_time,
                mod_date,
                e["crc"],
                e["compress_size"],
                e["file_size"],
                len(fname_bytes),
                len(extra),
                len(comment),
                0,                   # disk number start
                0,                   # internal attributes
                0,                   # external attributes
                e["local_offset"],
            ))
            self._buf.write(fname_bytes)
            self._buf.write(extra)
            self._buf.write(comment)

        cd_size = self._buf.tell() - cd_offset

        self._buf.write(b"PK\x05\x06")
        self._buf.write(_EOCD_FMT.pack(
            0,                       # disk number
            0,                       # disk where CD starts
            len(self._cd),           # entries on this disk
            len(self._cd),           # total entries
            cd_size,
            cd_offset,
            0,                       # comment length
        ))

        os.makedirs(os.path.dirname(self._path), exist_ok=True)
        payload = self._buf.getvalue()
        with open(self._path, "wb") as f:
            f.write(payload)
        print(f"[RawZipWriter] Written: {self._path}  ({len(payload):,} bytes)")

        # Verify it is a valid ZIP
        try:
            with zipfile.ZipFile(self._path, "r") as zv:
                entries = zv.namelist()
            print(f"[RawZipWriter] Verified OK: {len(entries)} entries: {entries}")
        except Exception as ex:
            print(f"[RawZipWriter] Verification FAILED: {ex}")


# ---------------------------------------------------------------------------
# Build test_d_rawcopy.hwpx
# ---------------------------------------------------------------------------

def build_rawcopy():
    print(f"\n{'='*60}")
    print("Building test_d_rawcopy.hwpx")
    print(f"  Original : {ORIGINAL_HWPX}")
    print(f"  Pass1    : {PASS1_HWPX}")
    print(f"  Output   : {OUTPUT_HWPX}")
    print(f"{'='*60}\n")

    # Pull the modified section0.xml from form_pass1.hwpx
    with zipfile.ZipFile(PASS1_HWPX, "r") as z:
        modified_section = z.read(SECTION_ENTRY)
    print(f"[INFO] Modified section0.xml uncompressed size: {len(modified_section):,} bytes")

    writer = RawZipWriter(OUTPUT_HWPX)

    with zipfile.ZipFile(ORIGINAL_HWPX, "r") as z_orig:
        for info in z_orig.infolist():
            if info.filename == SECTION_ENTRY:
                writer.add_data(info, modified_section,
                                compress_type=zipfile.ZIP_DEFLATED)
                print(f"[MODIFIED] {info.filename}  (re-compressed, flag_bits=0)")
            else:
                writer.add_raw(ORIGINAL_HWPX, info)
                print(f"[RAW COPY] {info.filename:<44} flag_bits={info.flag_bits}  compress_size={info.compress_size:,}")

    writer.close()


# ---------------------------------------------------------------------------
# Flag-bits comparison report
# ---------------------------------------------------------------------------

def print_comparison():
    files = [
        ("ORIGINAL", ORIGINAL_HWPX),
        ("PASS1",    PASS1_HWPX),
        ("RAWCOPY",  OUTPUT_HWPX),
    ]

    # Gather data: {label: {filename: ZipInfo}}
    data        = {}
    all_entries = []
    for label, path in files:
        data[label] = {}
        with zipfile.ZipFile(path, "r") as z:
            for info in z.infolist():
                data[label][info.filename] = info
                if info.filename not in all_entries:
                    all_entries.append(info.filename)

    print(f"\n{'='*100}")
    print("FLAG_BITS / COMPRESS_SIZE COMPARISON")
    print(f"{'='*100}")
    hdr = (f"{'Entry':<44}  {'ORIG':>6}  {'PASS1':>6}  {'RAW':>6}"
           f"  {'orig_csz':>10}  {'pass1_csz':>10}  {'raw_csz':>10}  note")
    print(hdr)
    print("-" * len(hdr))

    for fname in all_entries:
        orig  = data["ORIGINAL"].get(fname)
        pass1 = data["PASS1"].get(fname)
        raw   = data["RAWCOPY"].get(fname)

        def fb(x): return str(x.flag_bits)    if x else "---"
        def cs(x): return f"{x.compress_size:,}" if x else "---"

        note = ""
        if fname == SECTION_ENTRY:
            note = "<-- modified XML (flag_bits=0 expected)"
        elif orig and raw and orig.flag_bits != raw.flag_bits:
            note = "<-- MISMATCH (bug!)"
        elif orig and raw and orig.compress_size != raw.compress_size:
            note = "<-- compress_size differs (raw-copy bug?)"

        print(f"  {fname:<44}  {fb(orig):>6}  {fb(pass1):>6}  {fb(raw):>6}"
              f"  {cs(orig):>10}  {cs(pass1):>10}  {cs(raw):>10}  {note}")

    print(f"\nLEGEND:")
    print(f"  flag_bits=4  Enhanced deflating (original Hancom ZIP)")
    print(f"  flag_bits=0  Normal deflating   (Python zipfile.writestr default)")
    print(f"  ORIG column = form_to_fillout.hwpx (untouched original)")
    print(f"  PASS1 column = form_pass1.hwpx (current hwpx_editor.py output)")
    print(f"  RAW column  = test_d_rawcopy.hwpx (this script)")
    print()
    print(f"  For RAWCOPY, only '{SECTION_ENTRY}'")
    print(f"  should have flag_bits=0. All other entries should match ORIGINAL.")
    print()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    build_rawcopy()
    print_comparison()
