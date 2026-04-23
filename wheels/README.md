# Prebuilt Windows wheels

These wheels let a production Windows machine install the formatter's
dependencies without reaching PyPI (useful when corporate TLS inspection
aborts large PyPI downloads).

Targets CPython 3.12 on `win_amd64`. Regenerate from an internet-connected
machine whenever a dependency changes:

```bash
pip download --dest wheels/ --platform win_amd64 --python-version 3.12 \
             --only-binary=:all: python-docx lxml
```

Commit the result. On the offline box:

```powershell
.\.venv\Scripts\python.exe -m pip install --no-index `
    --find-links .\wheels python-docx lxml
.\.venv\Scripts\python.exe -m pip install --no-deps --no-build-isolation -e .
```
