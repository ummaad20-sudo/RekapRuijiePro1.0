from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.utils import platform

from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import os


class RekapApp(App):

    def build(self):
        root = BoxLayout(orientation='vertical', padding=10, spacing=10)

        title = Label(
            text="Rekap Data Ruijie Pro",
            size_hint=(1, 0.1),
            bold=True
        )
        root.add_widget(title)

        # area hasil
        self.result_label = Label(
            text="Silakan pilih file Excel...",
            size_hint_y=None,
            halign="left",
            valign="top"
        )
        self.result_label.bind(
            texture_size=lambda instance, value: setattr(instance, 'height', value[1])
        )

        scroll = ScrollView()
        scroll.add_widget(self.result_label)
        root.add_widget(scroll)

        # tombol bawah
        btn = Button(
            text="Pilih File Excel",
            size_hint=(1, 0.1)
        )
        btn.bind(on_press=self.buka_file)
        root.add_widget(btn)

        return root

    # ===============================
    # 📂 buka file chooser
    # ===============================
    def buka_file(self, instance):

        # tentukan path android
        start_path = self.get_android_path()

        chooser = FileChooserListView(
            path=start_path,
            filters=["*.xlsx"],  # 🔥 filter excel
        )

        popup = Popup(
            title="Pilih File Excel",
            content=chooser,
            size_hint=(0.95, 0.95)
        )

        chooser.bind(
            on_submit=lambda ch, sel, touch: self.proses_file(sel, popup)
        )

        popup.open()

    # ===============================
    # 📱 path android download
    # ===============================
    def get_android_path(self):
        if platform == "android":
            return "/storage/emulated/0/Download"
        return os.getcwd()

    # ===============================
    # 📊 proses excel
    # ===============================
    def proses_file(self, selection, popup):
        if not selection:
            return

        file_path = selection[0]
        popup.dismiss()

        try:
            wb = load_workbook(file_path)
            sheet = wb.active
        except Exception as e:
            self.result_label.text = f"Gagal membuka file:\n{e}"
            return

        rekap_detail = defaultdict(lambda: {"jumlah": 0, "total": 0})
        rekap_tanggal = defaultdict(int)

        header = [cell.value for cell in sheet[1]]

        try:
            kolom_grup = header.index("Grup pengguna")
            kolom_harga = header.index("Harga")
            kolom_tanggal = header.index("Diaktifkan di")
        except Exception:
            self.result_label.text = (
                "Header tidak sesuai!\n"
                "Pastikan ada:\n"
                "- Grup pengguna\n"
                "- Harga\n"
                "- Diaktifkan di"
            )
            return

        # ===============================
        # loop data
        # ===============================
        for row in sheet.iter_rows(min_row=2, values_only=True):
            grup = row[kolom_grup]
            harga = row[kolom_harga]
            tanggal_full = row[kolom_tanggal]

            if not (grup and harga and tanggal_full):
                continue

            # format tanggal
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

        # ===============================
        # format hasil
        # ===============================
        hasil = "===== HASIL REKAP =====\n\n"
        tanggal_terakhir = None

        for (tanggal, grup), data in sorted(rekap_detail.items()):

            if tanggal_terakhir and tanggal != tanggal_terakhir:
                hasil += f">>> TOTAL {tanggal_terakhir} : Rp {rekap_tanggal[tanggal_terakhir]:,}\n"
                hasil += "-----------------------------\n"

            if tanggal != tanggal_terakhir:
                hasil += f"\nTanggal : {tanggal}\n"

            hasil += f"  Grup   : {grup}\n"
            hasil += f"  Jumlah : {data['jumlah']}\n"
            hasil += f"  Total  : Rp {data['total']:,}\n\n"

            tanggal_terakhir = tanggal

        if tanggal_terakhir:
            hasil += f">>> TOTAL {tanggal_terakhir} : Rp {rekap_tanggal[tanggal_terakhir]:,}\n"

        self.result_label.text = hasil


if __name__ == "__main__":
    RekapApp().run()
