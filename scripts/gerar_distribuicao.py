"""Gera o executável portátil, ZIP de distribuição e SHA-256."""
from pathlib import Path
import hashlib
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from app_version import APP_VERSAO


def main():
    if sys.platform != "win32":
        raise SystemExit("A distribuição Windows deve ser gerada no Windows.")
    subprocess.run([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile",
        "--windowed", "--name", "Marine PGR", "--icon", str(ROOT / "resources/app.ico"),
        "--add-data", f"{ROOT / 'resources'};resources", "--paths", str(ROOT / "src"),
        "--collect-all", "tkinterdnd2", "--collect-data", "pdfminer",
        "--distpath", str(ROOT / "dist"), "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"), str(ROOT / "main.py"),
    ], cwd=ROOT, check=True)
    saida = ROOT / "distribuicao"
    saida.mkdir(exist_ok=True)
    exe = ROOT / "dist/Marine PGR.exe"
    zip_path = saida / f"Marine_PGR_v{APP_VERSAO}_Windows.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as pacote:
        pacote.write(exe, "Marine PGR.exe")
        pacote.writestr("COMO_USAR.txt", "Marine PGR\r\n\r\nExtraia o ZIP e abra Marine PGR.exe.\r\nNao precisa instalar Python.\r\nUse PDFs originais da seguradora, sempre RC-DC.\r\nConfira o PGR antes de enviar.\r\n")
    with zip_path.open("rb") as arquivo:
        hash_zip = hashlib.sha256()
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            hash_zip.update(bloco)
        digest = hash_zip.hexdigest()
    (saida / f"{zip_path.stem}.sha256").write_text(f"{digest}  {zip_path.name}\n", encoding="ascii")
    print(f"PRONTO: {zip_path}")


if __name__ == "__main__":
    main()
