"""Build a deterministic source bundle, excluding databases and historical binaries."""

from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
VERSION = "3.0.0"


def main():
    files = [
        ROOT / name
        for name in (
            "README.md",
            "LICENSE",
            "SECURITY.md",
            "CHANGELOG.md",
            "pyproject.toml",
            "requirements.txt",
            "matching.py",
            "workbench.py",
            "depositsmatcher.py",
            ".gitignore",
            ".gitattributes",
            "MANIFEST.in",
        )
    ]
    for directory in ("reconcile", "tests", "scripts", "examples", "docs", ".github"):
        files.extend(
            item
            for item in (ROOT / directory).rglob("*")
            if item.is_file()
            and item.suffix in {".py", ".md", ".csv", ".png", ".yml"}
            and "__pycache__" not in item.parts
        )
    destination = ROOT / "dist"
    destination.mkdir(exist_ok=True)
    archive = destination / f"deposits-matcher-{VERSION}-source.zip"
    with ZipFile(archive, "w") as bundle:
        for source in sorted(files):
            relative = source.relative_to(ROOT).as_posix()
            entry = ZipInfo(
                f"deposits-matcher-{VERSION}/{relative}", date_time=(1980, 1, 1, 0, 0, 0)
            )
            entry.compress_type = ZIP_DEFLATED
            entry.external_attr = 0o644 << 16
            bundle.writestr(entry, source.read_bytes(), compresslevel=9)
    checksum = sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix(".zip.sha256").write_text(f"{checksum}  {archive.name}\n", encoding="utf-8")
    print(f"{archive.name}: {checksum}")


if __name__ == "__main__":
    main()
