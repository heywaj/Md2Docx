# Md2Docx (Windows one-click Markdown to Word)

A lightweight GUI tool for converting Markdown files to Word (`.docx`) using `pandoc`.

## Features

- Chinese/English UI switch
- Modernized UI style (larger action buttons, cleaner card layout, improved readability)
- Two-tier interface: `Standard` (default) and `Advanced`
- Standard mode:
  - Paste mode: paste markdown text, set output dir and file name, convert directly
  - Single file mode: convert one markdown file
- Single file conversion: one Markdown file -> one Word file
- Folder batch conversion: convert all Markdown files in a folder
- Optional `reference.docx` template for custom Word style
- Runs as Python script or packaged `.exe`
- Packaged `.exe` can include bundled `pandoc.exe` (no runtime install needed)
- Startup argument support (for drag-and-drop onto `.exe`)

## Files

- `app.py`: main GUI app
- `build_exe.bat`: Windows build script for `PyInstaller`
- `requirements-build.txt`: build dependency list

## Initialize local Git repo

```bash
git init
git add .
git commit -m "init md2docx gui tool"
```

## Prerequisites

- Python 3.9+
- `pandoc` installed and available in `PATH` (for running from source or building exe)
- `PyInstaller` (for building exe)

For end users of `Md2Docx.exe`, no separate pandoc installation is required.

Check pandoc:

```bash
pandoc --version
```

## Run locally

```bash
python app.py
```

## Local test on macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m py_compile app.py
```

Prepare test markdown files:

```bash
mkdir -p testdata/sub
cat > testdata/a.md <<'EOF'
# A

hello **md**
EOF
cat > testdata/sub/b.markdown <<'EOF'
# B

- item 1
- item 2
EOF
```

Use GUI to verify:

1. Single mode: `testdata/a.md` -> output folder `test_out_single`
2. Folder mode (recursive): `testdata` -> output folder `test_out_batch`

You can also pass startup argument:

```bash
python app.py testdata/a.md
```

## Build Windows exe

1. Install build dependency:

```bash
pip install -r requirements-build.txt
```

2. Build:

```bat
build_exe.bat
```

`build_exe.bat` will find `pandoc.exe` from `PATH` and bundle it into the final exe.

3. Output:

- `dist\\Md2Docx.exe`

## GitHub Actions (build Windows exe on macOS dev workflow)

Workflow file:

- `.github/workflows/build-windows-exe.yml`
- `.github/workflows/release-windows-exe.yml`

Trigger options:

1. Push to `main`
2. Manual run (`workflow_dispatch`)

Result:

- Download artifact `Md2Docx-windows-exe` from Actions page

## Publish release with Windows exe

When you push a tag like `v1.0.0`, GitHub Actions will:

1. Build `Md2Docx.exe` on Windows runner
2. Create/update the GitHub Release for that tag
3. Upload `Md2Docx.exe` as release asset

Commands:

```bash
git tag v1.0.0
git push origin v1.0.0
```

## Usage

1. Launch app (`python app.py` or `Md2Docx.exe`)
2. Select language (`中文` / `English`)
3. Select interface edition:
   - `Standard` (default)
   - `Advanced`
4. In `Standard`:
   - `Paste Mode`: paste markdown text, choose output dir, enter output file name, click convert
   - `Single File Mode`: choose one markdown file and output dir, click convert
5. In `Advanced`:
   - choose `Single File` or `Folder Batch`
   - set input, output, optional template, then convert

### Drag-and-drop to exe

You can drag a `.md` file (or a folder) onto `Md2Docx.exe`.
The app will pre-fill input/output fields automatically.

## Notes

- In folder mode, output keeps the same relative subfolder structure.
- Supported Markdown extensions: `.md`, `.markdown`, `.mdown`, `.mkd`.
- If using source mode (`python app.py`) and pandoc is missing, add pandoc to `PATH`.
