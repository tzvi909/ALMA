"""Static checks of excel/vba/*.bas for common Excel VBA compile errors:
undeclared assignment targets / loop variables (Option Explicit), duplicate
public procedure names across modules, procedures over the 64 KB limit."""
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BAS = sorted(glob.glob(os.path.join(ROOT, "excel", "vba", "*.bas")))
PROC = re.compile(r"^(Public |Private )?(Sub|Function) (\w+)\((.*?)\)?(\s+As\s+\w+(\(\))?)?\s*(_)?$", re.I)


def strip(line):
    out, q = [], False
    for ch in line:
        if ch == '"':
            q = not q
            out.append(ch)
        elif ch == "'" and not q:
            break
        else:
            out.append("" if q else ch)
    return "".join(out)


def names_in_decl(text):
    text = re.sub(r"\([^()]*\)", "", text)          # array bounds
    out = []
    for part in text.split(","):
        part = re.sub(r"^\s*(Optional|ByVal|ByRef|ParamArray)\s+", "", part.strip(), flags=re.I)
        part = re.sub(r"^\s*(Optional|ByVal|ByRef|ParamArray)\s+", "", part, flags=re.I)
        m = re.match(r"(\w+)", part)
        if m:
            out.append(m.group(1).lower())
    return out


def main():
    module_names, public_procs, problems = set(), {}, []
    modules = {}
    for p in BAS:
        src = open(p, encoding="utf-8").read().replace("\r\n", "\n")
        lines = [strip(l) for l in src.split("\n")]
        joined, buf = [], ""
        for l in lines:                                # join continuation lines
            if l.rstrip().endswith(" _"):
                buf += l.rstrip()[:-1] + " "
            else:
                joined.append(buf + l)
                buf = ""
        modules[os.path.basename(p)] = joined
        for l in joined:
            m = re.match(r"^\s*(Public |Private )?(Sub|Function) (\w+)", l, re.I)
            if m:
                module_names.add(m.group(3).lower())
                if (m.group(1) or "Public ").strip().lower() == "public":
                    public_procs.setdefault(m.group(3).lower(), []).append(os.path.basename(p))
    for n, mods in public_procs.items():
        if len(mods) > 1:
            problems.append("duplicate public procedure %s in %s" % (n, mods))
    globals_ = set(module_names)
    for mod, lines in modules.items():
        for l in lines:
            if re.match(r"^\s*(Public|Private|Dim|Const|Public Const|Private Const)\s", l, re.I) and \
                    not re.match(r"^\s*(Public |Private )?(Sub|Function|Type)\b", l, re.I):
                d = re.sub(r"^\s*(Public|Private|Dim)\s+(Const\s+)?", "", l, flags=re.I)
                d = re.sub(r"=.*$", "", d)
                globals_.update(names_in_decl(d))
            m = re.match(r"^\s*(Public |Private )?Type (\w+)", l, re.I)
            if m:
                globals_.add(m.group(2).lower())
    for mod, lines in modules.items():
        local, proc, size = set(), None, 0
        for i, l in enumerate(lines, 1):
            m = re.match(r"^\s*(Public |Private )?(Sub|Function) (\w+)\s*\((.*)\)", l, re.I)
            if m:
                proc, local, size = m.group(3), set(names_in_decl(m.group(4))) | {m.group(3).lower()}, 0
                continue
            if proc is None:
                continue
            size += len(l)
            if re.match(r"^\s*End (Sub|Function)", l, re.I):
                if size > 60000:
                    problems.append("%s.%s is %d bytes (limit 64 KB)" % (mod, proc, size))
                proc = None
                continue
            for dm in re.finditer(r"\bDim\s+(.*)$", l, re.I):
                local.update(names_in_decl(dm.group(1)))
            for stmt in l.split(":"):
                stmt = stmt.strip()
                m2 = re.match(r"^(?:Set\s+|For\s+)?(\w+)(\([^=]*\))?\s*=", stmt, re.I)
                if m2 and not re.match(r"^(If|ElseIf|Do|Loop|While|Case|Select)\b", stmt, re.I):
                    name = m2.group(1).lower()
                    if name not in local and name not in globals_:
                        problems.append("%s:%d %s: '%s' assigned but not declared" % (mod, i, proc, m2.group(1)))
                m3 = re.match(r"^ReDim\s+(?:Preserve\s+)?(\w+)", stmt, re.I)
                if m3 and m3.group(1).lower() not in local and m3.group(1).lower() not in globals_:
                    problems.append("%s:%d %s: ReDim of undeclared '%s'" % (mod, i, proc, m3.group(1)))
    for p in problems:
        print("PROBLEM", p)
    print("%d modules, %d procedures, %d problems" % (len(modules), len(module_names), len(problems)))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
