"""A mappában lévő .doc/.docx/.odt fájlok fejezetcímeit Címsor 1 (Heading 1) stílusra állítja.

A .doc és .odt fájlokat a LibreOffice segítségével .docx-re konvertálja, mert a
python-docx csak .docx formátumot tud kezelni. Az eredmény mindig .odt.
"""

import argparse
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

BLACK = RGBColor(0, 0, 0)

# A szöveg elején és végén lévő, önálló fejezetnek számító részek
FRONT_BACK_MATTER = [
    "bevezetés", "bevezető", "előszó", "prológus", "prológ", "előhang", "ajánlás",
    "utószó", "epilógus", "epilóg", "utóhang", "zárszó", "befejezés",
    "függelék", "jegyzetek", "köszönetnyilvánítás", "a szerző megjegyzése",
]

# Pl. "1. FEJEZET", "12. fejezet", "Első fejezet", "Előszó", "Epilógus:"
CHAPTER_PATTERN = re.compile(
    r"^\s*(\d+\.\s*fejezet|[a-záéíóöőúüű]+\s+fejezet|"
    + "|".join(re.escape(name) for name in FRONT_BACK_MATTER)
    + r")\s*[.:]?\s*$",
    re.IGNORECASE,
)

# Csak számmal jelölt fejezetek: "1.", "2." vagy "I.", "II." (a pont kötelező, különben
# egy díszbetűs "I" is római egyesnek látszana)
ARABIC_NUMBER_PATTERN = re.compile(r"^\s*(\d{1,3})\.\s*$")
ROMAN_NUMBER_PATTERN = re.compile(r"^\s*([IVXLC]+)\.\s*$")
ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}


def roman_to_int(roman: str) -> int | None:
    """Római szám értéke, vagy None, ha nem szabályos (pl. "IIII", "VX")."""
    total = 0
    for i, char in enumerate(roman):
        value = ROMAN_VALUES[char]
        if i + 1 < len(roman) and value < ROMAN_VALUES[roman[i + 1]]:
            total -= value
        else:
            total += value
    return total if int_to_roman(total) == roman else None


def int_to_roman(number: int) -> str:
    result = ""
    for value, letters in ((100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"),
                           (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        count, number = divmod(number, value)
        result += letters * count
    return result


def longest_numbered_sequence(candidates: list[tuple[int, int]]) -> list[int]:
    """A leghosszabb 1, 2, 3, ... sorozat a (bekezdés index, szám) jelöltekből.

    A sorozat 1-gyel kezdődik, és dokumentumsorrendben mindig a következő számmal
    folytatódik; a közben lévő, oda nem illő számokat kihagyja. Legalább két elem kell.
    """
    best: list[int] = []
    for start, (_, number) in enumerate(candidates):
        if number != 1:
            continue
        sequence = []
        expected = 1
        for index, value in candidates[start:]:
            if value == expected:
                sequence.append(index)
                expected += 1
        if len(sequence) > len(best):
            best = sequence
    return best if len(best) >= 2 else []


def numbered_chapter_indices(paragraphs) -> set[int]:
    """A csak számmal jelölt fejezetcímek bekezdés-indexei (arab és római külön sorozat)."""
    arabic, roman = [], []
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text
        if match := ARABIC_NUMBER_PATTERN.match(text):
            arabic.append((index, int(match.group(1))))
        elif (match := ROMAN_NUMBER_PATTERN.match(text)) and (value := roman_to_int(match.group(1))):
            roman.append((index, value))
    return set(longest_numbered_sequence(arabic)) | set(longest_numbered_sequence(roman))


# A python-docx csak .docx-et kezel, a többit előbb LibreOffice-szal konvertáljuk
CONVERTED_INPUTS = (".doc", ".odt")
SUPPORTED_INPUTS = (".docx", *CONVERTED_INPUTS)

FLATPAK_ID = "org.libreoffice.LibreOffice"

# Egy konvertálás legfeljebb ennyi másodpercig tarthat (egyes fájlokon a LibreOffice végtelen ciklusba kerül)
CONVERT_TIMEOUT = 180


def find_libreoffice() -> list[str]:
    """A LibreOffice indítóparancsa.

    Előbb a Flatpak, aztán a rendszerszintű telepítés.
    """
    flatpak = shutil.which("flatpak")
    if flatpak and subprocess.run([flatpak, "info", FLATPAK_ID], capture_output=True).returncode == 0:
        return [flatpak, "run", FLATPAK_ID]
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        sys.exit("Hiba: a LibreOffice nem található, .doc fájlhoz szükséges.")
    return [soffice]


def kill_processes_with_arg(text: str) -> None:
    """Leállítja azokat a folyamatokat, amelyek parancssorában szerepel a szöveg."""
    for cmdline_file in Path("/proc").glob("[0-9]*/cmdline"):
        try:
            if text in cmdline_file.read_bytes().decode(errors="replace"):
                os.kill(int(cmdline_file.parent.name), signal.SIGKILL)
        except (OSError, ValueError):
            pass  # közben kilépett, vagy nem a miénk


def convert_with_libreoffice(src: Path, target_format: str, out_dir: Path) -> Path:
    """Konvertálás LibreOffice-szal (pl. doc -> docx, docx -> doc).

    Minden konvertálás saját, ideiglenes LibreOffice-profilt kap. Így nem akad össze
    egy éppen megnyitott LibreOffice-szal, és ha egy fájlon a LibreOffice elakad,
    pontosan a hozzá tartozó folyamatok állíthatók le.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    # A profil a kimeneti mappában van, mert a Flatpak a gép /tmp mappáját nem látja
    with tempfile.TemporaryDirectory(dir=out_dir, prefix=".lo-profil-") as profile:
        profile_url = Path(profile).as_uri()
        try:
            proc = subprocess.run(
                [*find_libreoffice(), f"-env:UserInstallation={profile_url}",
                 "--headless", "--convert-to", target_format,
                 "--outdir", str(out_dir), str(src)],
                capture_output=True,
                text=True,
                timeout=CONVERT_TIMEOUT,
            )
        except subprocess.TimeoutExpired:
            # A Flatpak sandboxában futó LibreOffice a leállított indító után is tovább futna
            kill_processes_with_arg(profile_url)
            raise RuntimeError(
                f"a LibreOffice {CONVERT_TIMEOUT} mp alatt sem végzett a konvertálással "
                f"({src.name} -> {target_format}), leállítva"
            ) from None
    result = out_dir / f"{src.stem}.{target_format}"
    if proc.returncode != 0 or not result.exists():
        hint = ""
        if "no export filter" in proc.stderr + proc.stdout:
            hint = "\nTipp: telepítsd a Writer modult: sudo dnf install libreoffice-writer"
        raise RuntimeError(f"a konvertálás nem sikerült: {src}\n{proc.stderr.strip()}{hint}")
    return result


def get_heading1_style(document):
    """A Heading 1 stílus lekérése (vagy létrehozása), középre igazítva, fekete színnel."""
    try:
        style = document.styles["Heading 1"]
    except KeyError:
        style = document.styles.add_style("Heading 1", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = document.styles["Normal"]
        style.font.bold = True
        style.font.size = Pt(16)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.font.color.rgb = BLACK
    return style


def format_chapters(docx_path: Path, output_path: Path) -> list[str]:
    """A fejezetcím-bekezdéseket Heading 1 stílusra állítja és elmenti."""
    document = Document(str(docx_path))
    heading1 = get_heading1_style(document)

    paragraphs = document.paragraphs
    numbered = numbered_chapter_indices(paragraphs)

    found = []
    for index, paragraph in enumerate(paragraphs):
        # A már címsorként formázott részek (pl. alfejezetként egy Utóhang) szintjét meghagyjuk
        # A stílus None, ha a bekezdés a dokumentumban nem létező stílusra hivatkozik
        style = paragraph.style
        already_heading = (
            style is not None and style.name.startswith("Heading") and style != heading1
        )
        is_chapter = index in numbered or CHAPTER_PATTERN.match(paragraph.text)
        if is_chapter and not already_heading:
            paragraph.style = heading1
            # A bekezdésen lévő közvetlen formázás felülírná a stílust
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.color.rgb = BLACK
            found.append(paragraph.text.strip())

    document.save(str(output_path))
    return found


def is_valid_odt(path: Path) -> bool:
    """Ellenőrzi, hogy a fájl ép ODT-e (ZIP, benne a szöveget tartalmazó content.xml)."""
    try:
        with zipfile.ZipFile(path) as archive:
            return "content.xml" in archive.namelist() and archive.testzip() is None
    except (OSError, zipfile.BadZipFile):
        return False


def process_file(path: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    # A Flatpak saját /tmp-t lát, ezért az átmeneti mappa a kimeneti mappában van
    with tempfile.TemporaryDirectory(dir=output_dir) as tmp:
        tmp_dir = Path(tmp)
        source = path
        if path.suffix.lower() in CONVERTED_INPUTS:
            source = convert_with_libreoffice(path, "docx", tmp_dir / "be")
        formatted = tmp_dir / f"{path.stem}.docx"
        found = format_chapters(source, formatted)
        output_odt = convert_with_libreoffice(formatted, "odt", output_dir)

    print(f"{path.name}: {len(found)} fejezetcím formázva")
    for title in found:
        print(f"  - {title}")
    print(f"  Mentve: {output_odt}")
    return output_odt


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_dir", nargs="?", default="../test", type=Path,
                        help="A .doc/.docx/.odt fájlokat tartalmazó mappa (alapértelmezett: ../test)")
    parser.add_argument("-o", "--output-dir", type=Path, default=None,
                        help="Kimeneti mappa (alapértelmezett: <input_dir>/formazott)")
    parser.add_argument("--megtart", action="store_true",
                        help="Sikeres átalakítás után se törölje az eredeti fájlt")
    args = parser.parse_args()

    input_dir: Path = args.input_dir.resolve()
    output_dir: Path = (args.output_dir or input_dir / "formazott").resolve()

    files = sorted(
        p for p in input_dir.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_INPUTS
        and not p.name.startswith(("~$", ".~lock"))
    )
    if not files:
        sys.exit(f"Nincs .doc/.docx/.odt fájl a mappában: {input_dir}")

    failed = []
    written: set[Path] = set()
    for path in files:
        try:
            # Azonos nevű bemenetek (pl. konyv.doc és konyv.odt) felülírnák egymás kimenetét
            if output_dir / f"{path.stem}.odt" in written:
                raise RuntimeError("ugyanilyen nevű kimenet már készült ebben a futásban, kihagyva")
            output_odt = process_file(path, output_dir)
            written.add(output_odt)
        except Exception as error:  # egy hibás fájl ne állítsa le a többit
            print(f"{path.name}: HIBA – {type(error).__name__}: {error}", file=sys.stderr)
            failed.append(path.name)
            continue

        if args.megtart:
            continue
        if is_valid_odt(output_odt):
            path.unlink()
            print(f"  Eredeti törölve: {path}")
        else:
            print(f"{path.name}: HIBA – a kimenet nem ép ODT, az eredeti megmaradt", file=sys.stderr)
            failed.append(path.name)

    print(f"\nKész: {len(files) - len(failed)}/{len(files)} fájl feldolgozva.")
    if failed:
        print(f"Sikertelen ({len(failed)}):", file=sys.stderr)
        for name in failed:
            print(f"  - {name}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
