import smtplib
import time
import sys
import os
import shutil
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import pandas as pd
import PyPDF2
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer


class ReportGenerator:
    SMTP_SERVER = "mail.dashte.com"
    SMTP_PORT = 465
    SENDER_EMAIL = "demo_for_dip_report@dashtekengineers.com"
    SENDER_PASSWORD = "GSi%%?s,X%pF"
    EMAIL_DATA_FILE = "email_data.xlsx"
    INPUT_FILE = "HYOSUNG_Dispatch_Details.xlsx"
    OUTPUT_FILE = "output_file.xlsx"
    INDEX_PDF_FOLDER = "index_pdf"
    ATTACHMENT_FOLDER = "Attachment_folder"
    INPUT_FOLDER = "input_folder"

    def __init__(self):
        self.cd_path = os.getcwd()
        self.input_file_path = os.path.join(self.cd_path, self.INPUT_FILE)
        self.output_file_path = os.path.join(self.cd_path, self.OUTPUT_FILE)
        self.index_folder = os.path.join(self.cd_path, self.INDEX_PDF_FOLDER)
        self.destination_directory = os.path.join(self.cd_path, self.ATTACHMENT_FOLDER)
        self.input_folder = os.path.join(self.cd_path, self.INPUT_FOLDER)

    def sort_the_file_as_per_part_no(self):
        df = pd.read_excel(self.input_file_path)
        grouped = df.groupby('Part no')

        with pd.ExcelWriter(self.output_file_path, engine='xlsxwriter') as writer:
            for invoice_no, group in grouped:
                cleaned_invoice_no = ''.join(c for c in str(invoice_no) if c.isalnum() or c in ['-', '_', ' '])
                if cleaned_invoice_no and cleaned_invoice_no not in writer.sheets:
                    group.to_excel(writer, sheet_name=f'Part_no_{cleaned_invoice_no}', index=False)

        print(f"Output Excel file has been generated:----Done---- {self.output_file_path}")

    def make_pdf_file_from_the_output_excel(self):
        output_folder = os.path.join(self.cd_path, "index_pdf")
        os.makedirs(output_folder, exist_ok=True)
        input_file_path = os.path.join(self.cd_path, "output_file.xlsx")
        xls = pd.ExcelFile(input_file_path)
        font_size = 28

        for sheet_name in xls.sheet_names:
            data = xls.parse(sheet_name)
            pdf_file_name = f"{data.iloc[0]['Part no']}.pdf"
            pdf_file_path = os.path.join(output_folder, pdf_file_name)

            self.create_pdf_from_data(data, pdf_file_path, font_size)

            print(f"Output PDF file has been generated:----Done---- {pdf_file_path}")

    def create_pdf_from_data(self, data, pdf_file_path, font_size):
        pdf_buffer = BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4))
        elements = []

        data_list = [data.columns.tolist()]
        data_list.extend(data.values.tolist())
        col_widths = [100] * len(data.columns)
        row_heights = [24] * len(data_list)

        header_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), (1, 1, 1)),
            ('TEXTCOLOR', (0, 0), (-1, 0), (0, 0, 0)),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold', font_size),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, 0), 1, (0, 0, 0)),
        ])

        data_style = TableStyle([
            ('BACKGROUND', (0, 1), (-1, -1), (1, 1, 1)),
            ('TEXTCOLOR', (0, 1), (-1, -1), (0, 0, 0)),
            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica', font_size),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 12),
            ('GRID', (0, 1), (-1, -1), 1, (0, 0, 0)),
        ])

        table = Table(data_list, colWidths=col_widths, rowHeights=row_heights)
        table.setStyle([
            ('BACKGROUND', (0, 0), (-1, 0), (1, 1, 1)),
            ('GRID', (0, 0), (-1, -1), 1, (0, 0, 0)),
        ])

        for i in range(len(data_list)):
            if i == 0:
                table.setStyle(header_style)
            else:
                table.setStyle(data_style)

        available_height = doc.height - (doc.topMargin + doc.bottomMargin)
        table_height = table.wrap(doc.width, available_height)[1]

        if table_height < available_height:
            space_height = (available_height - table_height) / 2
            elements.append(Spacer(1, space_height))

        elements.append(table)
        doc.build(elements)

        pdf_buffer.seek(0)

        with open(pdf_file_path, 'wb') as f:
            f.write(pdf_buffer.read())

    def place_index_file_in_associated_folder(self):
        for pdf_file in os.listdir(self.index_folder):
            if pdf_file.lower().endswith(".pdf"):
                pdf_file_path = os.path.join(self.index_folder, pdf_file)
                folder_name = os.path.splitext(pdf_file)[0]
                destination_folder = os.path.join(self.input_folder, folder_name)

                if os.path.exists(destination_folder):
                    new_pdf_path = os.path.join(destination_folder, "1.pdf")
                    shutil.move(pdf_file_path, new_pdf_path)
                else:
                    print(f"Folder '{folder_name}' not found in '{self.input_folder}'. Skipping '{pdf_file}'.")

        print("Index file generated ----Done---- ")

    def merge_pdfs_in_folder(self, folder_path):
        pdf_files = [os.path.join(folder_path, filename) for filename in os.listdir(folder_path) if
                     filename.lower().endswith(".pdf") and filename[0].isalpha() and 'A' <= filename[0].upper() <= 'Z']

        if not pdf_files:
            return
        pdf_files = sorted(pdf_files, key=lambda x: self.get_numeric_prefix(x))
        pdf_merger = PyPDF2.PdfWriter()
        for pdf_file in pdf_files:
            try:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                for page_num in range(len(pdf_reader.pages)):
                    pdf_merger.add_page(pdf_reader.pages[page_num])
            except Exception as e:
                print(f"Skipping '{pdf_file}' due to an error: {e}")

        if len(pdf_merger.pages) > 0:
            folder_name = os.path.basename(folder_path)
            output_pdf = os.path.join(folder_path, f"{folder_name}_merged.pdf")
            temp_output_pdf = os.path.join(folder_path, f"{folder_name}_temp_merged.pdf")

            with open(temp_output_pdf, 'wb') as output_file:
                pdf_merger.write(output_file)

            shutil.move(temp_output_pdf, output_pdf)

            print(f"Merged PDFs in '{folder_name}' sorted from 'A' to 'Z' and saved as '{folder_name}_merged.pdf'.")

    def get_numeric_prefix(self, filename):
        # Extract the numeric prefix from the filename
        parts = os.path.basename(filename).split('.')
        if len(parts) > 0 and parts[0].isdigit():
            numeric_prefix = ''.join(c for c in parts[0] if c.isdigit())
            return int(numeric_prefix)
        else:
            print("No numeric prefix found.")
            return float('inf')

    def search_file_by_prefix(self, start_path, part_no):
        for root, dirs, files in os.walk(start_path):
            for file in files:
                if part_no in file and file.endswith("_merged.pdf"):
                    return os.path.join(root, file)

    def send_emails_from_excel(self, part_no, subject_of_mail):
        try:
            df = pd.read_excel(os.path.join(self.cd_path, self.EMAIL_DATA_FILE))
            to_email = "edshfsdgfdgfdhjg@gmail.com"
            sheet_prefix = "Part_no_"
            sheet_name = sheet_prefix + part_no
            table_html = self.display_sheet_data(sheet_name)
            attachment_path = self.search_file_by_prefix(os.path.join(self.cd_path, "Attachment_folder/"), part_no)

            if attachment_path:
                print(f"File found at: {attachment_path}")
                self.send_email_with_attachment(self.SENDER_EMAIL, to_email, subject_of_mail, self.SMTP_SERVER,
                                                self.SMTP_PORT, self.SENDER_PASSWORD, attachment_path, table_html)
            else:
                print(f"No matching file with prefix '{part_no}' found.")

        except pd.errors.ParserError:
            print(f"Error: Unable to read '{self.EMAIL_DATA_FILE}'.")

    def display_sheet_data(self, sheet_name):
        input_file_path = os.path.join(os.getcwd(), "output_file.xlsx")
        xls = pd.ExcelFile(input_file_path)

        try:
            if sheet_name in xls.sheet_names:
                data = xls.parse(sheet_name)
                table_html = data.to_html(index=False)
                return table_html
            else:
                print(f"Sheet '{sheet_name}' not found in the Excel file.")
                return None
        except pd.errors.ParserError:
            print(f"Error: Sheet '{sheet_name}' could not be parsed.")
            return None

    def send_email_with_attachment(self, sender_email, to_email, subject_of_mail, smtp_server, smtp_port,
                                   sender_password, attachment_path, table_html):
        cd_path = os.getcwd()
        input_file_path = os.path.join(cd_path, "email_data.xlsx")
        df = pd.read_excel(input_file_path)

        for index, row in df.iterrows():
            message_1 = row['Message_1']
            message_2 = row['Message_2']
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = to_email
            message = f"{message_1}\n\n\n\n{table_html}\n\n\n\n{message_2}"
            msg['Subject'] = subject_of_mail

            msg.attach(MIMEText(message, 'html'))

            if attachment_path:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(open(attachment_path, 'rb').read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
                msg.attach(part)

            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, to_email, msg.as_string())

            print("Email sent successfully.")

    def copy_merged_pdfs(self, source_path, destination_path):
        if not os.path.exists(destination_path):
            os.makedirs(destination_path)

        for item in os.listdir(destination_path):
            item_path = os.path.join(destination_path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)

        for root, dirs, files in os.walk(source_path):
            for file in files:
                if file.endswith("_merged.pdf"):
                    source_file_path = os.path.join(root, file)
                    destination_file_path = os.path.join(destination_path, file)

                    shutil.copy2(source_file_path, destination_file_path)

    def clean_and_trim(self, s):
        s = s.strip()
        s = ''.join(c for c in s if c.isalnum() or c.isspace())
        return s

    def create_email_subject(self, file_path):
        if os.path.exists(file_path):
            xls = pd.ExcelFile(file_path)
            sheet_count = len(xls.sheet_names)
            for i, sheet_name in enumerate(xls.sheet_names, start=1):
                print("email data ===>", i)
                part_name = pd.read_excel(file_path, sheet_name).iloc[0]['Part name']
                part_no = pd.read_excel(file_path, sheet_name).iloc[0]['Part no']
                concatenated_result = f"{part_name}-{part_no}"
                subject_of_mail = f"PDI Report for {concatenated_result} Mail-{i}/{sheet_count}"
                self.send_emails_from_excel(part_no, subject_of_mail)
        else:
            print(f"File '{file_path}' not found.")

    def loading_spinner(self, progress):
        bar_length = 20
        block = int(round(bar_length * progress))
        progress_str = "#" * block + "-" * (bar_length - block)
        sys.stdout.write('\r[' + progress_str + f'] {int(progress * 100)}%')
        sys.stdout.flush()

    def simulate_process(self, duration_seconds):
        start_time = time.time()
        elapsed_time = 0

        while elapsed_time < duration_seconds:
            progress = elapsed_time / duration_seconds
            self.loading_spinner(progress)
            time.sleep(1)
            elapsed_time = time.time() - start_time

if __name__ == "__main__":
    report_generator = ReportGenerator()
    report_generator.sort_the_file_as_per_part_no()
    report_generator.make_pdf_file_from_the_output_excel()
    report_generator.place_index_file_in_associated_folder()
    file_path = os.path.join(os.getcwd(), "output_file.xlsx")

    for folder_name in os.listdir(report_generator.input_folder):
        folder_path = os.path.join(report_generator.input_folder, folder_name)

        if os.path.isdir(folder_path):
            report_generator.merge_pdfs_in_folder(folder_path)

    print("PDF merging completed.----Done---- ")
    report_generator.copy_merged_pdfs(report_generator.input_folder, report_generator.destination_directory)
    print(f"Processing please wait")

    try:
        report_generator.simulate_process(60)
    except KeyboardInterrupt:
        print("\nOperation canceled by user.")
    print("\n")
    ip = str(input("Do you want to send these reports on mail. Y/N : "))

    if ip.upper() == "Y":
        try:
            report_generator.create_email_subject(file_path)
        except Exception as e:
            print(f"Error: {e}")


#-----
