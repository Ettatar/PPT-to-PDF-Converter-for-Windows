import comtypes.client
import os
import sys
import traceback
from pathlib import Path

PDF_FORMAT = 32


def log(msg):
    print(msg)


def get_ppt_files(folder):
    folder = Path(folder)

    files = []
    for f in folder.iterdir():
        if f.name.startswith("~$"):
            continue

        if f.suffix.lower() in [".ppt", ".pptx"]:
            files.append(f)

    return files


def start_powerpoint():
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    return powerpoint


def convert_one(powerpoint, input_path, output_path):

    try:
        presentation = powerpoint.Presentations.Open(
            str(input_path),
            WithWindow=False
        )

        presentation.SaveAs(
            str(output_path),
            PDF_FORMAT
        )

        presentation.Close()

        log(f"[OK] {input_path.name}")

    except Exception as e:
        log(f"[ERROR] {input_path.name}")
        traceback.print_exc()


def batch_convert(folder):

    folder = Path(folder).resolve()

    files = get_ppt_files(folder)

    if not files:
        log("No PPT files found")
        return

    log(f"Found {len(files)} files")

    powerpoint = None

    try:
        powerpoint = start_powerpoint()

        for file in files:

            pdf_path = file.with_suffix(".pdf")

            if pdf_path.exists():
                log(f"[SKIP] {pdf_path.name}")
                continue

            convert_one(
                powerpoint,
                file,
                pdf_path
            )

    finally:
        if powerpoint:
            powerpoint.Quit()
            log("PowerPoint closed")


def main():

    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = os.getcwd()

    batch_convert(folder)


if __name__ == "__main__":
    main()
