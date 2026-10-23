# Fejezet formázó

Az eredmény mindig `.odt` fájl.

A megadott mappában lévő `.doc` / `.docx` / `.odt` fájlokban megkeresi a fejezetcímeket
(pl. `1. FEJEZET`, `Prológus`, `Epilógus`), és **Címsor 1** (Heading 1) stílusra állítja őket.

A `.doc` és `.odt` fájlokat a LibreOffice konvertálja `.docx`-re (ehhez telepítve kell
lennie; a Flatpakos változatot is felismeri).

## Telepítés

```bash
python3 -m venv .venv
.venv/bin/pip install -e .
```

## Használat

```bash
.venv/bin/fejezet-formazo ../test             # eredmény: ../test/formazott/*.odt
.venv/bin/fejezet-formazo ../test --megtart   # az eredeti fájlokat nem törli
```

Sikeres átalakítás után az eredeti fájl **törlődik** (kivéve `--megtart` esetén).
Csak akkor, ha a kimenet ép `.odt`; sikertelen átalakításnál az eredeti megmarad.

### `.venv` nélkül

A rendszerre telepített `python3`-mal is működik, ha a `python-docx` csomag elérhető
(pl. `sudo apt install python3-docx` vagy `pip install --user python-docx`).
A projekt mappájából futtatva:

```bash
python3 -m fejezet_formazo.main ../test    # eredmény: ../test/formazott/*.odt
```
