#!/usr/bin/env python3
import sys
from pathlib import Path
import pdfplumber

def main():
    folder = Path.cwd()
    out_path = folder / "pdf_words_dump.txt"  # один общий файл для всех PDF
    pdf_files = sorted([p for p in folder.glob("*.pdf")] + [p for p in folder.glob("*.PDF")])

    if not pdf_files:
        print("В этой папке PDF не найдено.")
        return

    with out_path.open("w", encoding="utf-8", newline="") as f:
        f.write("file\tpage\tx0\ty0\tx1\ty1\ttext\n")

        for pdf_path in pdf_files:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    for pageno, page in enumerate(pdf.pages, start=1):
                        # извлекаем слова с координатами
                        words = page.extract_words(
                            use_text_flow=True,
                            keep_blank_chars=False,
                            extra_attrs=["x0", "x1", "top", "bottom"]
                        )
                        # сортировка: сверху-вниз, слева-направо
                        words.sort(key=lambda w: (w["top"], w["x0"]))

                        for w in words:
                            txt = (w.get("text") or "").replace("\t", " ").replace("\r", " ").replace("\n", " ")
                            f.write(f"{pdf_path.name}\t{pageno}\t{w['x0']:.2f}\t{w['top']:.2f}\t{w['x1']:.2f}\t{w['bottom']:.2f}\t{txt}\n")
            except Exception as e:
                # если какой-то PDF не читается, зафиксируем и пойдём дальше
                with out_path.open("a", encoding="utf-8") as ferr:
                    ferr.write(f"{pdf_path.name}\tERROR\t0\t0\t0\t0\t{e}\n")

    print(f"Готово. Результат: {out_path}")

if __name__ == "__main__":
    main()