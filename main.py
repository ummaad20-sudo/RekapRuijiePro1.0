import os
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from plyer import filechooser
from openpyxl import load_workbook
from collections import defaultdict


class RekapApp(App):

    def build(self):
        root = BoxLayout(orientation="vertical", padding=10, spacing=10)

        # ===== JUDUL DI ATAS =====
        title = Label(
            text="REKAP PENJUALAN EXCEL",
            size_hint=(1, 0.1),
            color=(0, 0, 0, 1)
        )
        root.add_widget(title)

        # ===== AREA HASIL (TENGAH & BESAR) =====
        scroll = ScrollView(size_hint=(1, 0.8))

        self.output = TextInput(
            readonly=True,
            size_hint_y=None,
            text="Hasil rekap akan tampil di sini...",
            foreground_color=(0, 0, 0, 1),
            background_color=(1, 1, 1, 1)
        )

        self.output.bind(minimum_height=self.output.setter('height'))
        scroll.add_widget(self.output)
        root.add_widget(scroll)

        # ===== TOMBOL DI PALING BAWAH =====
        btn = Button(
            text="Pilih File Excel",
            size_hint=(1, 0.1),
            background_color=(0.2, 0.6, 1, 1)
        )
        btn.bind(on_press=self.buka_file)
        root.add_widget(btn)

        return root

    def buka_file(self, instance):
        filechooser.open_file(
            title="Pilih File Excel",
            filters=[("Excel Files", "*.xlsx")],
            on_selection=self.file_selected
        )

    def file_selected(self, selection):
        if selection:
            self.proses_excel(selection[0])

    def proses_excel(self, file_path):
        try:
            wb = load_workbook(file_path)
            sheet = wb.active

            rekap = defaultdict(int)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                tanggal = row[11]  # Kolom L
                harga = row[4]     # Kolom E

                if tanggal and harga:
                    rekap[str(tanggal)[:10]] += int(harga)

            hasil = "=== REKAP PER TANGGAL ===\n\n"

            for tgl, total in sorted(rekap.items()):
                hasil += f"{tgl} : Rp {total:,}\n"

            self.output.text = hasil

        except Exception as e:
            self.output.text = f"Terjadi kesalahan:\n{e}"


if __name__ == "__main__":
    RekapApp().run()
