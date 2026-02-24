from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime

class RekapApp(App):

    def build(self):
        self.layout = BoxLayout(orientation='vertical')

        self.label = Label(text="Rekap Data Ruijie Pro", size_hint=(1, 0.1))
        self.layout.add_widget(self.label)

        self.result_label = Label(text="", size_hint_y=None)
        self.result_label.bind(texture_size=self.result_label.setter('size'))

        scroll = ScrollView()
        scroll.add_widget(self.result_label)

        self.layout.add_widget(scroll)

        btn = Button(text="Pilih File Excel", size_hint=(1, 0.1))
        btn.bind(on_press=self.buka_file)
        self.layout.add_widget(btn)

        return self.layout

    def buka_file(self, instance):
        content = FileChooserListView()
        popup = Popup(title="Pilih File Excel",
                      content=content,
                      size_hint=(0.9, 0.9))

        content.bind(on_submit=lambda x, selection, touch: self.proses_file(selection, popup))
        popup.open()

    def proses_file(self, selection, popup):
        if not selection:
            return

        file_path = selection[0]
        popup.dismiss()

        try:
            wb = load_workbook(file_path)
            sheet = wb.active
        except:
            self.result_label.text = "Gagal membuka file!"
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
