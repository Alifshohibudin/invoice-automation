import pandas as pd
import re
import os
from docxtpl import DocxTemplate

# path custom folder
output_dir = r"C:\Users\PT DIRCA\Documents\PT DIRGANTARA CAHAYA ABADI\Tukar Faktur\invoices baru" #---simpan di lokal folder kalian
os.makedirs(output_dir, exist_ok=True)   # buat folder kalau belum ada

def safe_filename(text):
    # ganti karakter terlarang dengan "-"
    return re.sub(r'[\/:*?"<>|]', '-', str(text))

def idr(n):
    return "Rp {:,}".format(int(n)).replace(",", ".")

def terbilang(n):
    angka = ["", "Satu", "Dua", "Tiga", "Empat", "Lima",
             "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    def _tb(x):
        if x < 12:
            return angka[x]
        elif x < 20:
            return _tb(x - 10) + " Belas"
        elif x < 100:
            return _tb(x // 10) + " Puluh" + (" " + _tb(x % 10) if x % 10 else "")
        elif x < 200:
            return "Seratus " + _tb(x - 100)
        elif x < 1000:
            return _tb(x // 100) + " Ratus" + (" " + _tb(x % 100) if x % 100 else "")
        elif x < 2000:
            return "Seribu " + _tb(x - 1000)
        elif x < 1000000:
            return _tb(x // 1000) + " Ribu" + (" " + _tb(x % 1000) if x % 1000 else "")
        elif x < 1000000000:
            return _tb(x // 1000000) + " Juta" + (" " + _tb(x % 1000000) if x % 1000000 else "")
        else:
            return str(x)
    return _tb(int(n)).strip()

def generate_invoice(excel_path, template_path):
    headers = pd.read_excel(excel_path, sheet_name="Header").to_dict(orient="records")
    items_all = pd.read_excel(excel_path, sheet_name="Items").to_dict(orient="records")
    patients_all = pd.read_excel(excel_path, sheet_name="Patients").to_dict(orient="records")

    for header in headers:
        no_invoice = header["no_invoice"]
        items = [i for i in items_all if i["no_invoice"] == no_invoice]
        patients = [p for p in patients_all if p["no_invoice"] == no_invoice]

        for idx, item in enumerate(items, start=1):
            item["no"] = idx  # nomor urut bulat 1,2,3...


        for item in items:
            item["jml"] = int(item["jml"])
            item["value_num"] = item["jml"] * item["harga"]
            item["harga"] = idr(item["harga"])
            item["value"] = idr(item["value_num"])

        subtotal = sum(i["value_num"] for i in items)
        ppn = int(subtotal * header.get("ppn_persen", 11) / 100)
        total = subtotal + ppn

        context = {
            "no_invoice": header["no_invoice"],
            "tgl_invoice": header["tgl_invoice"],
            "kepada": header["kepada"],
            "alamat": header["alamat"],
            "items": items,
            "subtotal": idr(subtotal),
            "ppn": idr(ppn),
            "ppn_persen": header.get("ppn_persen", 11),
            "total": idr(total),
            "terbilang": terbilang(total) + " Rupiah",
            "patients": patients,
            "judul_perusahaan": header["judul_perusahaan"],
            "ttd_nama": header["ttd_nama"],
            "ttd_jabatan": header["ttd_jabatan"]
        }

        doc = DocxTemplate(template_path)
        doc.render(context)

        safe_no_invoice = safe_filename(no_invoice)
        output_path = os.path.join(output_dir, f"invoice_{safe_no_invoice}.docx")
        doc.save(output_path)
        print(f"✅ Invoice berhasil dibuat: {output_path}")

if __name__ == "__main__":
    excel_path = "invoice_input.xlsx"
    template_path = "template_invoice.docx"
    generate_invoice(excel_path, template_path)
