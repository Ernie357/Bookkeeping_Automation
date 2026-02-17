from QuickbooksInvoiceHandler import QuickbooksInvoiceHandler
from QRCodeHandler import QRCodeHandler
from ExcelHandler import ExcelHandler
from ExcelHandler import CorrespondingData
from MailMergeHandler import MailMergeHandler
import traceback

realm_id = "9341456354221278"
access_token = "eyJhbGciOiJkaXIiLCJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwieC5vcmciOiJIMCJ9..9Jfb4e02f9r6NgcIFhJi6w.CLHQPvxir-rnrwJz60E1DXyVRrKEmq9ulgzTptuHCDMf-rCN-qPQqLUFqD5Q4FDzv-luPfngJOoO3Na3qXvmcPzQecEocHQ9-0G1UeL1e5Pj_E7tqXTUrWZZSNumie-PhJznhF6vJU2Vsr4zZ9mDS1-nGDFPm1XYyKVH9u37g-AhEXIkoCj5gBy0dHaRvQmdfDjmLy0LxowbztPkuJ58Kbxx2rhkDXXy_fwKR4XHfpJioZqvCieDjl6GmvV-y5h0auOuqUv76_Ok97Wkg8sXgGOcodbEP6Cshqeu7RVsZruNNzASerHoNyNeinzwGSEx9PgdyUwM1srhH0Br4qWSs3-crX8JbVoLuLV4UOULO0yftHMEm_bFlN1o_ABrYMu8WEHNX1kTFvDceNfxnMoas_89k9x9BQR_eYdg34yAisAel1S6d5ZMJpnlvpDp0YLboEOK2cjOFNnAIGChzKhrbc8HSRCxuOBkJAzAOaNrsio.wTPbOUtZ3vEPhy7IaUNR9g"
if __name__ == "__main__":
    try:
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