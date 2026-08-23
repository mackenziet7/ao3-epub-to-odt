from pathlib import Path

def suggest_odt_path(epub: str, output_folder: str) -> str:
    """
    Builds a non-colliding output path for the ODT file.
    e.g. given epub="my_fic.epub" and folder="C:/output",
    returns "C:/output/my_fic_book.odt", or
            "C:/output/my_fic_book_2.odt" if that already exists, etc.
    """
    base = Path(output_folder) / (Path(epub).stem + "_book.odt")
    if not base.exists():
        return str(base)
    counter = 2
    original_stem = base.stem
    while base.exists():
        base = base.with_stem(original_stem + f"_{counter}")
        counter += 1
    return str(base)