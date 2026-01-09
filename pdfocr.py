
from pathlib import Path
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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


def rename_to_txt(md_file_path):
    """Rename markdown file to txt file."""
    txt_file_path = md_file_path.with_suffix('.txt')
    md_file_path.rename(txt_file_path)
    return txt_file_path


def send_email_with_attachment(file_path, recipient_email, smtp_server="smtp.163.com", smtp_port=465,
                                sender_email=None, sender_password=None):
    """
    Send an email with the file as an attachment.

    Args:
        file_path: Path to the file to attach
        recipient_email: Email address of the recipient
        smtp_server: SMTP server address (default: smtp.163.com)
        smtp_port: SMTP port (default: 465 for SSL)
        sender_email: Sender's email address
        sender_password: Sender's email password or app-specific password
    """
    if not sender_email or not sender_password:
        raise ValueError("sender_email and sender_password must be provided")

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"PDF OCR Output - {file_path.stem}"

    body = f"Please find attached the OCR output file: {file_path.name}"
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {file_path.name}')
    msg.attach(part)

    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")
        raise


def main():
    import os

    os.environ['HTTP_PROXY'] = 'http://proxy.net.sap.corp:8080'
    os.environ['HTTPS_PROXY'] = 'http://proxy.net.sap.corp:8080'

    data_folder = Path(__file__).parent / "../../tests/data"
    input_doc_path = Path(os.getenv("INPUT_PDF_PATH", "./data/吕思勉全集 7 隋唐五代史 上 (吕思勉著).pdf"))

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

    # Rename to .txt file
    txt_file = rename_to_txt(output_file)
    print(f"File renamed to: {txt_file}")

    # Send email with attachment
    # Email credentials are read from environment variables
    # Set these in your environment or .env file:
    # - SENDER_EMAIL: Your email address
    # - SENDER_PASSWORD: Your email password or app-specific password
    # - RECIPIENT_EMAIL: Recipient's email address
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    recipient_email = os.getenv("RECIPIENT_EMAIL")

    if sender_email and sender_password and recipient_email:
        try:
            send_email_with_attachment(
                file_path=txt_file,
                recipient_email=recipient_email,
                sender_email=sender_email,
                sender_password=sender_password
            )
        except Exception as e:
            print(f"Email sending failed: {str(e)}")
    else:
        print("Email credentials not found in environment variables. Skipping email sending.")
        print("Set SENDER_EMAIL, SENDER_PASSWORD, and RECIPIENT_EMAIL to enable email functionality.")


if __name__ == "__main__":
    main()