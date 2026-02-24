from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView

from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime

# Android imports
from android import activity
from jnius import autoclass
import tempfile


class RekapApp(App):

    def build(self):
        self.layout = BoxLayout(orientation='vertical')

        self.label = Label(
            text="Rekap Data Ruijie Pro",
            size_hint=(1, 0.1)
        )
        self.layout.add_widget(self.label)

        self.result_label = Label(
            text="",
            size_hint_y=None
        )
        self.result_label.bind(
            texture_size=self.result_label.setter('size')
        )

        scroll = ScrollView()
        scroll.add_widget(self.result_label)
        self.layout.add_widget(scroll)

        btn = Button(
            text="Pilih File Excel",
            size_hint=(1, 0.1)
        )
        btn.bind(on_press=self.buka_file)
        self.layout.add_widget(btn)

        return self.layout

    # =========================
    # ANDROID FILE PICKER (SAF)
    # =========================
    def buka_file(self, instance):
        Intent = autoclass('android.content.Intent')
        intent = Intent(Intent.ACTION_GET_CONTENT)
        intent.setType("*/*")
        intent.addCategory(Intent.CATEGORY_OPENABLE)

        activity.bind(on_activity_result=self.on_file_selected)
        activity.startActivityForResult(intent, 1)

    def on_file_selected(self, request_code, result_code, intent):
        if request_code == 1 and result_code == -1:
            uri = intent.getData()
            if uri:
                self.baca_file_dari_uri(uri)

    def baca_file_dari_uri(self, uri):
        PythonActivity = autoclass('org.kivy.android.PythonActivity')
        context = PythonActivity.mActivity
        resolver = context.getContentResolver()
        input_stream = resolver.openInputStream(uri)

        temp_file = tempfile.NamedTemporaryFile(delete=False)
        temp_file.write(input_stream.read())
        temp_file.close()

        self.proses_file(temp_file.name)

    # =========================
    # PROSES EXCEL
    # =========================
    def proses_file(self, file_path):

        try:
            wb = load_workbook(file_path)
            sheet = wb.active
        except Exception as e:
            self.result_label.text = f"Gagal membuka file!\n{str(e)}"
            return

        rekap_detail = defaultdict(lambda: {"jumlah": 0, "total": 0})
        rekap_tanggal = defaultdict(int)

        header = [cell.value for cell in sheet[1]]

        try:
            kolom_grup = header.index("Grup pengguna")
            kolom_harga = header.index("Harga")
            kolom_tanggal = header.index("Diaktifkan di")
        except:
            self.result_label.text = "Header tidak sesuai!"
            return

        for row in sheet.iter_rows(min_row=2, values_only=True):
            grup = row[kolom_grup]
            harga = row[kolom_harga]
            tanggal_full = row[kolom_tanggal]

            if grup and harga and tanggal_full:

                if isinstance(tanggal_full, datetime):
                    tanggal = tanggal_full.strftime("%Y/%m/%d")
                else:
                    tanggal = str(tanggal_full).split(" ")[0]

                try:
                    harga_int = int(harga)
                except:
                    continue

                key = (tanggal, grup)
                rekap_detail[key]["jumlah"] += 1
                rekap_detail[key]["total"] += harga_int
                rekap_tanggal[tanggal] += harga_int

        hasil = "===== HASIL REKAP =====\n\n"
        tanggal_terakhir = None

        for (tanggal, grup), data in sorted(rekap_detail.items()):

            if tanggal_terakhir and tanggal != tanggal_terakhir:
                hasil += f">>> TOTAL {tanggal_terakhir} : Rp {rekap_tanggal[tanggal_terakhir]}\n"
                hasil += "-----------------------------\n"

            if tanggal != tanggal_terakhir:
                hasil += f"\nTanggal : {tanggal}\n"

            hasil += f"  Grup   : {grup}\n"
            hasil += f"  Jumlah : {data['jumlah']}\n"
            hasil += f"  Total  : Rp {data['total']}\n\n"

            tanggal_terakhir = tanggal

        if tanggal_terakhir:
            hasil += f">>> TOTAL {tanggal_terakhir} : Rp {rekap_tanggal[tanggal_terakhir]}\n"

        self.result_label.text = hasil


if __name__ == "__main__":
    RekapApp().run()
