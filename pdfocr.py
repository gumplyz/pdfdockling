
from pathlib import Path
from datetime import datetime

from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import (
    PdfPipelineOptions,
    TableStructureOptions,
    TesseractCliOcrOptions,
    EasyOcrOptions,
    OcrMacOptions
)
from docling.document_converter import DocumentConverter, PdfFormatOption


def create_output_folder():
    """Create a timestamped output folder and return the path."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = Path("./output") / timestamp
    output_folder.mkdir(parents=True, exist_ok=True)
    return output_folder


def main():
    import os

    os.environ['HTTP_PROXY'] = 'http://proxy.net.sap.corp:8080'
    os.environ['HTTPS_PROXY'] = 'http://proxy.net.sap.corp:8080'

    data_folder = Path(__file__).parent / "../../tests/data"
    input_doc_path = Path("./data/在李克农身边的日子.pdf")

    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_ocr = True
    pipeline_options.do_table_structure = True
    pipeline_options.table_structure_options = TableStructureOptions(
        do_cell_matching=True
    )

    # Any of the OCR options can be used: EasyOcrOptions, TesseractOcrOptions,
    # TesseractCliOcrOptions, OcrMacOptions (macOS only), RapidOcrOptions
    # ocr_options = EasyOcrOptions(force_full_page_ocr=True)
    # ocr_options = TesseractOcrOptions(force_full_page_ocr=True)
    # ocr_options = OcrMacOptions(force_full_page_ocr=True)
    # ocr_options = RapidOcrOptions(force_full_page_ocr=True)
    # ocr_options = TesseractOcrOptions(force_full_page_ocr=True, lang=["chi_sim","eng"])
    ocr_options = OcrMacOptions(force_full_page_ocr=True, lang=["zh-Hans", "en-US"])
    pipeline_options.ocr_options = ocr_options

    converter = DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_options,
            )
        }
    )

    doc = converter.convert(input_doc_path).document
    md = doc.export_to_markdown()

    # Create output folder and save markdown
    output_folder = create_output_folder()
    output_file = output_folder / f"{input_doc_path.stem}.md"
    output_file.write_text(md, encoding='utf-8')

    print(f"Markdown output saved to: {output_file}")
    print(f"\nPreview:\n{md[:500]}...")  # Print first 500 characters as preview


if __name__ == "__main__":
    main()