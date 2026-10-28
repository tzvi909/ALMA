"""Build a VBA project binary (vbaProject.bin) from .bas source files.

Implements the parts of [MS-OVBA] (compression, dir / PROJECT / PROJECTwm
streams, CMG/DPB/GC encryption) and [MS-CFB] (compound file container) needed
for a *source-only* project: the _VBA_PROJECT stream holds only its 7-byte
header and no compiled p-code, so Office recompiles the modules from source
when the workbook is opened (the approach EPPlus uses to create VBA projects).
"""
import random
import struct
import uuid

THISWORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
WORKSHEET_BASE = "0{00020820-0000-0000-C000-000000000046}"
REFERENCES = [
    ("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"),
    ("Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared"
               "\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library"),
]


# ------------------------------------------------------------ MS-OVBA 2.4.1 compression
def _bitcount(diff):
    bc = 0
    while (1 << bc) < diff:
        bc += 1
    return max(bc, 4)


def _compress_chunk(buf, start, end):
    out = bytearray()
    table = {}
    pos = start

    def index(i):
        if i + 3 <= end:
            table.setdefault(bytes(buf[i:i + 3]), []).append(i)

    while pos < end:
        flag_at = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= end:
                break
            diff = pos - start
            best_len = best_off = 0
            if diff > 0 and pos + 3 <= end:
                bc = _bitcount(diff)
                limit = min((0xFFFF >> bc) + 3, end - pos)
                for c in reversed(table.get(bytes(buf[pos:pos + 3]), [])[-512:]):
                    n = 0
                    while n < limit and buf[c + n] == buf[pos + n]:
                        n += 1
                    if n > best_len:
                        best_len, best_off = n, pos - c
                        if n == limit:
                            break
            if best_len >= 3:
                out += struct.pack("<H", ((best_off - 1) << (16 - bc)) | (best_len - 3))
                flags |= 1 << bit
                for i in range(pos, pos + best_len):
                    index(i)
                pos += best_len
            else:
                out.append(buf[pos])
                index(pos)
                pos += 1
        out[flag_at] = flags
    return bytes(out)


def compress(data):
    data = bytes(data)
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        end = min(start + 4096, len(data))
        chunk = _compress_chunk(data, start, end)
        if len(chunk) <= 4096:
            out += struct.pack("<H", (len(chunk) + 2 - 3) | (0b011 << 12) | (1 << 15)) + chunk
        elif end - start == 4096:
            out += struct.pack("<H", 4095 | (0b011 << 12)) + data[start:end]
        else:
            raise ValueError("incompressible final chunk")
    return bytes(out)


def decompress(data):
    if not data or data[0] != 1:
        raise ValueError("bad compressed container signature")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack("<H", data[pos:pos + 2])[0]
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0b111 != 0b011:
            raise ValueError("bad chunk signature")
        chunk_end = pos + size
        pos += 2
        dstart = len(out)
        if not header & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flags & (1 << bit):
                    token = struct.unpack("<H", data[pos:pos + 2])[0]
                    pos += 2
                    bc = _bitcount(len(out) - dstart)
                    length = (token & (0xFFFF >> bc)) + 3
                    offset = (token >> (16 - bc)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# ------------------------------------------------------------------ PROJECT encryption
def encrypt(project_id, data, seed=None):
    """[MS-OVBA] 2.4.3.2 data encryption (used for CMG, DPB and GC)."""
    seed = random.randrange(256) if seed is None else seed
    key = sum(project_id.encode("ascii")) & 0xFF
    out = [seed, 2 ^ seed, key ^ seed]
    plain = bytes([seed]) * ((seed & 6) // 2) + struct.pack("<I", len(data)) + bytes(data)
    enc1, enc2, unenc1 = out[2], out[1], key
    for b in plain:
        e = b ^ ((enc2 + unenc1) & 0xFF)
        out.append(e)
        enc2, enc1, unenc1 = enc1, e, b
    return "".join("%02X" % x for x in out)


def decrypt(project_id, text):
    d = bytes.fromhex(text)
    seed, version, key = d[0], d[1] ^ d[0], d[2] ^ d[0]
    assert version == 2 and key == sum(project_id.encode("ascii")) & 0xFF
    enc1, enc2, unenc1 = d[2], d[1], key
    plain = bytearray()
    for e in d[3:]:
        b = e ^ ((enc2 + unenc1) & 0xFF)
        plain.append(b)
        enc2, enc1, unenc1 = enc1, e, b
    ignored = (seed & 6) // 2
    n = struct.unpack("<I", plain[ignored:ignored + 4])[0]
    return bytes(plain[ignored + 4:ignored + 4 + n])


# ------------------------------------------------------------------------ streams
def _rec(rid, data=b""):
    return struct.pack("<HI", rid, len(data)) + data


def dir_stream(project_name, modules, codepage=1252, lcid=0x409):
    enc = lambda s: s.encode("cp1252")  # noqa: E731
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))                 # PROJECTSYSKIND (Win32)
    d += _rec(0x0002, struct.pack("<I", lcid))              # PROJECTLCID
    d += _rec(0x0014, struct.pack("<I", lcid))              # PROJECTLCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", codepage))          # PROJECTCODEPAGE
    d += _rec(0x0004, enc(project_name))                    # PROJECTNAME
    d += _rec(0x0005) + _rec(0x0040)                        # PROJECTDOCSTRING (+ unicode)
    d += _rec(0x0006) + _rec(0x003D)                        # PROJECTHELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))                 # PROJECTHELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))                 # PROJECTLIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1425420523, 6)     # PROJECTVERSION
    d += _rec(0x000C) + _rec(0x003C)                        # PROJECTCONSTANTS (+ unicode)
    for name, libid in REFERENCES:
        d += _rec(0x0016, enc(name)) + _rec(0x003E, name.encode("utf-16-le"))   # REFERENCENAME
        lib = enc(libid)
        d += struct.pack("<HII", 0x000D, 4 + len(lib) + 6, len(lib)) + lib + struct.pack("<IH", 0, 0)
    d += _rec(0x000F, struct.pack("<H", len(modules)))      # PROJECTMODULES
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))            # PROJECTCOOKIE
    for m in modules:
        n = m["name"]
        d += _rec(0x0019, enc(n))                           # MODULENAME
        d += _rec(0x0047, n.encode("utf-16-le"))            # MODULENAMEUNICODE
        d += _rec(0x001A, enc(n)) + _rec(0x0032, n.encode("utf-16-le"))  # MODULESTREAMNAME
        d += _rec(0x001C) + _rec(0x0048)                    # MODULEDOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))             # MODULEOFFSET: source starts at 0
        d += _rec(0x001E, struct.pack("<I", 0))             # MODULEHELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))        # MODULECOOKIE
        d += struct.pack("<HI", 0x0021 if m["type"] == "module" else 0x0022, 0)  # MODULETYPE
        d += struct.pack("<HI", 0x002B, 0)                  # module terminator
    d += struct.pack("<HI", 0x0010, 0)                      # dir terminator
    return compress(bytes(d))


def project_stream(project_id, modules, project_name="VBAProject"):
    lines = ['ID="%s"' % project_id]
    for m in modules:
        lines.append(("Document=%s/&H00000000" if m["type"] == "document" else "Module=%s") % m["name"])
    lines += ['Name="%s"' % project_name, 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              'CMG="%s"' % encrypt(project_id, struct.pack("<I", 0)),
              'DPB="%s"' % encrypt(project_id, b"\x00"),
              'GC="%s"' % encrypt(project_id, b"\xff"),
              "", "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
              "", "[Workspace]"]
    lines += ["%s=0, 0, 0, 0, C" % m["name"] for m in modules]
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def projectwm_stream(modules):
    b = bytearray()
    for m in modules:
        b += m["name"].encode("cp1252") + b"\x00" + m["name"].encode("utf-16-le") + b"\x00\x00"
    return bytes(b + b"\x00\x00")


def document_source(name, base):
    return ('Attribute VB_Name = "%s"\r\nAttribute VB_Base = "%s"\r\nAttribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n' % (name, base))


# ------------------------------------------------------------------- MS-CFB writer
ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


def _cfb_key(name):
    return (len(name), name.upper())


def _tree(children, entries):
    """Balanced binary tree of sibling entries (CFB order); returns root id.
    Nodes on the deepest level of a non-perfect tree are red, all others black,
    which satisfies the red-black rules."""
    order = sorted(children, key=lambda i: _cfb_key(entries[i]["name"]))
    depth = {}

    def build(lo, hi, d):
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        node = order[mid]
        depth[node] = d
        entries[node]["left"] = build(lo, mid - 1, d + 1)
        entries[node]["right"] = build(mid + 1, hi, d + 1)
        return node

    root = build(0, len(order) - 1, 0)
    if order:
        dmax = max(depth.values())
        perfect = len(order) == 2 ** (dmax + 1) - 1
        for node, d in depth.items():
            entries[node]["color"] = 0 if (d == dmax and not perfect) else 1
    return root


def cfb(streams):
    """streams: {"path/name": bytes}; storages are created from the paths."""
    entries = [{"name": "Root Entry", "type": 5, "children": [], "data": b""}]
    storages = {"": 0}
    for path in sorted(streams):
        parts = path.split("/")
        parent = 0
        for k in range(len(parts) - 1):
            key = "/".join(parts[:k + 1])
            if key not in storages:
                storages[key] = len(entries)
                entries.append({"name": parts[k], "type": 1, "children": [], "data": b""})
                entries[parent]["children"].append(storages[key])
            parent = storages[key]
        entries.append({"name": parts[-1], "type": 2, "children": [], "data": bytes(streams[path])})
        entries[parent]["children"].append(len(entries) - 1)
    for e in entries:
        e.setdefault("left", NOSTREAM)
        e.setdefault("right", NOSTREAM)
        e.setdefault("color", 1)
    for e in entries:
        e["child"] = _tree(e["children"], entries) if e["children"] else NOSTREAM
    entries[0]["color"] = 1

    # mini stream for streams < 4096 bytes
    mini, minifat = bytearray(), []
    big = []
    for i, e in enumerate(entries):
        if e["type"] != 2:
            continue
        n = len(e["data"])
        if n < 4096:
            first = len(mini) // 64
            nsec = max(1, (n + 63) // 64)
            e["start"], e["size"] = first, n
            mini += e["data"] + b"\x00" * (nsec * 64 - n)
            minifat += [first + k + 1 for k in range(nsec - 1)] + [ENDOFCHAIN]
        else:
            big.append(i)
    n_dir = (len(entries) + 3) // 4
    n_minifat = (len(minifat) + 127) // 128
    n_ministream = (len(mini) + 511) // 512
    n_big = sum((len(entries[i]["data"]) + 511) // 512 for i in big)
    n_fat = 1
    while 128 * n_fat < n_fat + n_dir + n_minifat + n_ministream + n_big:
        n_fat += 1
    if n_fat > 109:
        raise ValueError("file too large for a header-only DIFAT")

    fat = [FREESECT] * (128 * n_fat)
    sectors = []
    nxt = 0

    def alloc(nsec, payload):
        nonlocal nxt
        first = nxt
        for k in range(nsec):
            fat[first + k] = first + k + 1 if k < nsec - 1 else ENDOFCHAIN
        payload = payload + b"\x00" * (nsec * 512 - len(payload))
        sectors.append((first, payload))
        nxt += nsec
        return first

    for k in range(n_fat):
        fat[k] = FATSECT
    nxt = n_fat
    dir_start_placeholder = nxt
    nxt += n_dir
    for k in range(n_dir):
        fat[dir_start_placeholder + k] = dir_start_placeholder + k + 1 if k < n_dir - 1 else ENDOFCHAIN
    minifat_start = alloc(n_minifat, b"".join(struct.pack("<I", x) for x in minifat)
                          + struct.pack("<I", FREESECT) * (n_minifat * 128 - len(minifat))) if n_minifat else ENDOFCHAIN
    ministream_start = alloc(n_ministream, bytes(mini)) if n_ministream else ENDOFCHAIN
    entries[0]["start"], entries[0]["size"] = ministream_start, len(mini)
    for i in big:
        d = entries[i]["data"]
        entries[i]["start"] = alloc((len(d) + 511) // 512, d)
        entries[i]["size"] = len(d)

    def dir_entry(e):
        name = e["name"].encode("utf-16-le") + b"\x00\x00"
        b = name + b"\x00" * (64 - len(name))
        b += struct.pack("<HBB", len(name), e["type"], e["color"])
        b += struct.pack("<III", e["left"], e["right"], e["child"])
        b += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        b += struct.pack("<IQ", e.get("start", 0) if e["type"] != 1 else 0, e.get("size", 0) if e["type"] != 1 else 0)
        return b

    empty = (b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
             + b"\x00" * 36 + struct.pack("<IQ", 0, 0))
    dirbytes = b"".join(dir_entry(e) for e in entries) + empty * (n_dir * 4 - len(entries))
    sectors.append((dir_start_placeholder, dirbytes))
    fatbytes = b"".join(struct.pack("<I", x) for x in fat)
    sectors.append((0, fatbytes))

    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", header, 40, 0, n_fat, dir_start_placeholder, 0, 4096,
                     minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = [k for k in range(n_fat)] + [FREESECT] * (109 - n_fat)
    struct.pack_into("<109I", header, 76, *difat)
    body = bytearray(512 * nxt)
    for first, payload in sectors:
        body[first * 512:first * 512 + len(payload)] = payload
    return bytes(header) + bytes(body)


# --------------------------------------------------------------------- project
def read_bas(path):
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    first = text.split("\r\n", 1)[0]
    if not first.startswith('Attribute VB_Name = "'):
        raise ValueError("%s must start with an Attribute VB_Name line" % path)
    return first.split('"')[1], text


def build(bas_files, sheet_codenames, project_name="VBAProject", project_id=None):
    """Return vbaProject.bin bytes with ThisWorkbook, one document module per
    worksheet code name and the given standard modules."""
    project_id = project_id or "{%s}" % str(uuid.uuid4()).upper()
    modules = [{"name": "ThisWorkbook", "type": "document",
                "source": document_source("ThisWorkbook", THISWORKBOOK_BASE)}]
    for cn in sheet_codenames:
        modules.append({"name": cn, "type": "document", "source": document_source(cn, WORKSHEET_BASE)})
    for path in bas_files:
        name, src = read_bas(path)
        modules.append({"name": name, "type": "module", "source": src})
    streams = {
        "PROJECT": project_stream(project_id, modules, project_name),
        "PROJECTwm": projectwm_stream(modules),
        "VBA/_VBA_PROJECT": bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
        "VBA/dir": dir_stream(project_name, modules),
    }
    for m in modules:
        streams["VBA/" + m["name"]] = compress(m["source"].encode("cp1252"))
    return cfb(streams), modules
