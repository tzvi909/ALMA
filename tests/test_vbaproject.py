"""Checks of the vbaProject.bin writer against independent readers (olefile, oletools)."""
import glob
import os
import random
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "excel"))
import vbaproject as vp  # noqa: E402

BAS = sorted(glob.glob(os.path.join(os.path.dirname(HERE), "excel", "vba", "*.bas")))


def test_compression_roundtrip():
    rnd = random.Random(1)
    samples = [b"", b"a", b"abcabcabcabc" * 500, bytes(rnd.randrange(256) for _ in range(9000)),
               b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"]
    samples += [open(p, "rb").read() for p in BAS]
    for s in samples:
        assert vp.decompress(vp.compress(s)) == s


def test_compression_matches_oletools():
    from oletools.olevba import decompress_stream
    for p in BAS:
        src = open(p, "rb").read()
        assert decompress_stream(bytearray(vp.compress(src))) == src


def test_spec_example():
    # input string of the [MS-OVBA] 3.2 compression example
    s = b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"
    assert vp.decompress(vp.compress(s)) == s


def test_encryption_roundtrip():
    pid = "{12345678-ABCD-4EF0-9876-0123456789AB}"
    for seed in range(256):
        for data in (b"\x00\x00\x00\x00", b"\x00", b"\xff"):
            assert vp.decrypt(pid, vp.encrypt(pid, data, seed)) == data


def test_project_readable_by_olevba(tmp_path=None):
    import olefile
    from oletools.olevba import VBA_Parser
    data, modules = vp.build(BAS, ["Sheet1", "Sheet2"])
    os.makedirs(os.environ.get("ALMA_TMP", os.path.join(os.path.dirname(HERE), "test_output")), exist_ok=True)
    out = os.path.join(os.environ.get("ALMA_TMP", os.path.join(os.path.dirname(HERE), "test_output")), "test_vbaProject.bin")
    with open(out, "wb") as fh:
        fh.write(data)
    ole = olefile.OleFileIO(out)
    names = sorted("/".join(e) for e in ole.listdir())
    assert "VBA/dir" in names and "PROJECT" in names and "VBA/ALMA_Main" in names
    vba = VBA_Parser(out)
    got = {os.path.splitext(name)[0]: code for (_, _, name, code) in vba.extract_macros()}
    for p in BAS:
        name, src = vp.read_bas(p)
        assert name in got, name
        assert got[name].replace("\r\n", "\n").strip() == src.replace("\r\n", "\n").strip()
    assert {"ThisWorkbook", "Sheet1", "Sheet2"} <= set(got)


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f()
            print("ok", k)
