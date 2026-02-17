from QuickbooksInvoiceHandler import QuickbooksInvoiceHandler
from QRCodeHandler import QRCodeHandler
from ExcelHandler import ExcelHandler
from ExcelHandler import CorrespondingData
from MailMergeHandler import MailMergeHandler
from dotenv import load_dotenv
import os
import traceback

if __name__ == "__main__":
    try:
        load_dotenv()
        realm_id = os.getenv('REALM_ID')
        access_token = os.getenv('ACCESS_TOKEN')
        if realm_id is None or access_token is None:
            raise Exception("Missing .env variables.")
        print("\n")
        qh = QuickbooksInvoiceHandler(realm_id=realm_id, access_token=access_token)
        qr = QRCodeHandler()
        excel = ExcelHandler(
            filename="Legal_Notices_2025_Master_List_V2.2.xlsx",
            worksheet_name="Bill&Cert"
        )
        print("\n")
        mm = MailMergeHandler(template_filename="Invoice_Template_For_Mail.docx")
        print("\n")
        qh.import_csv("QBO_Import_Data.csv")
        print(f"\nInvoice IDs: {qh.invoice_ids}")
        print(f"\nInvoice Numbers: {qh.invoice_numbers}")
        print("\n")
        qr.generate_qr_codes(target_dir="qr_codes", ids=qh.invoice_ids)
        qr_path_data_to_add = CorrespondingData(col_name="X", data=qr.code_paths)
        qr_link_data_to_add = CorrespondingData(col_name="Y", data=qr.code_links)
        merge_data_list = qr.add_qrs_excel(
            excel=excel,
            invoice_num_col_name="G",
            invoice_nums=qh.invoice_numbers,
            data_lists=[qr_path_data_to_add, qr_link_data_to_add]
        )
        mm.merge_multiple(merge_data_list=merge_data_list)
        print("\n")
        mm.close()
        print("\n")
    except:
        with open("errors.txt", "w") as logfile:
            traceback.print_exc(file=logfile)
        raise