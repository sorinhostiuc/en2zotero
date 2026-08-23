import json
import os
import zipfile


BASE = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE, "manifest.json"), "r", encoding="utf-8") as f:
    version = json.load(f)["version"]

xpi_name = f"en2zotero-{version}.xpi"
xpi_path = os.path.join(BASE, xpi_name)

runtime_roots = {"content", "locale"}
runtime_files = {"manifest.json", "bootstrap.js", "prefs.js"}
excluded_runtime_files = {
    "content/odf-writer.js",
    "content/ui/dialog-test.xhtml",
}

entries = []

for root, dirs, files in os.walk(BASE):
    rel_root = os.path.relpath(root, BASE)
    if rel_root == ".":
        dirs[:] = [d for d in dirs if d in runtime_roots]
    else:
        top = rel_root.split(os.sep, 1)[0]
        if top not in runtime_roots:
            dirs[:] = []
            continue

    for name in files:
        rel = os.path.relpath(os.path.join(root, name), BASE)
        rel_posix = rel.replace(os.sep, "/")
        top = rel.split(os.sep, 1)[0]
        if rel_posix in excluded_runtime_files:
            continue
        if rel_root == "." and name not in runtime_files:
            continue
        if rel_root != "." and top not in runtime_roots:
            continue
        entries.append((rel_posix, os.path.join(root, name)))

entries.sort(
    key=lambda entry: (
        0 if entry[0] == "manifest.json" else 1 if entry[0] in {"bootstrap.js", "prefs.js"} else 2,
        entry[0],
    )
)

with zipfile.ZipFile(xpi_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for arc, full in entries:
        zf.write(full, arc)

print(xpi_path)
