from pathlib import Path

import fitz


def extract_text_from_pdf(pdf_path: Path) -> str:
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo PDF: {pdf_path}")

    full_text_parts: list[str] = []

    with fitz.open(pdf_path) as doc:
        for page in doc:
            text = page.get_text()
            if text:
                full_text_parts.append(text)

    return "\n".join(full_text_parts)


def main() -> None:
    ruta = input("Ruta completa del PDF a leer: ").strip().strip('"')
    if not ruta:
        print("No se ingresó ruta de archivo.")
        return

    try:
        texto = extract_text_from_pdf(Path(ruta))
        print("\n===== TEXTO EXTRAÍDO DEL PDF =====\n")
        print(texto)
    except Exception as exc:
        print(f"Error al leer el PDF: {exc}")


if __name__ == "__main__":
    main()