"""Assemble src/part*.py + the preset geometry blob into the single-file pptx_to_html.py."""
import glob, os
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
parts = sorted(glob.glob(os.path.join(root, "src", "part*.py")))
out = "".join(open(p, encoding="utf-8").read() for p in parts)
blob = open(os.path.join(root, "src", "presets.b64")).read().strip()
out += '\n# ECMA-376 preset shape geometry (zlib + base64 JSON), see tools/build_presets.py\n_PRESET_B64 = """\n%s\n"""\n\nif __name__ == "__main__":\n    main()\n' % blob
open(os.path.join(root, "pptx_to_html.py"), "w", encoding="utf-8").write(out)
print(len(out.splitlines()), "lines")
