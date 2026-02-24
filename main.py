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
        self.layout = BoxLayout(orientation="vertical", padding=10, spacing=10)

        self.label = Label(text="Rekap Excel", size_hint=(1, 0.1))
        self.layout.add_widget(self.label)

        self.btn = Button(text="Pilih File Excel", size_hint=(1, 0.1))
        self.btn.bind(on_press=self.buka_file)
        self.layout.add_widget(self.btn)

        self.output = TextInput(readonly=True)
        scroll = ScrollView()
        scroll.add_widget(self.output)
        self.layout.add_widget(scroll)

        return self.layout

    def buka_file(self, instance):
        try:
            filechooser.open_file(
                title="Pilih File Excel",
                filters=[("Excel Files", "*.xlsx")],
                on_selection=self.file_selected
            )
        except Exception as e:
            self.output.text = f"Gagal membuka file picker:\n{e}"

    def file_selected(self, selection):
        if selection:
            file_path = selection[0]
            self.proses_excel(file_path)

    def proses_excel(self, file_path):
        try:
            wb = load_workbook(file_path)
            sheet = wb.active

            rekap = defaultdict(int)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                tanggal = row[11]  # Kolom L (index 11)
                harga = row[4]     # Kolom E (index 4)

                if tanggal and harga:
                    rekap[str(tanggal)[:10]] += int(harga)

            hasil = "=== REKAP PER TANGGAL ===\n\n"
            for tgl, total in rekap.items():
                hasil += f"{tgl} : Rp {total:,}\n"

            self.output.text = hasil

        except Exception as e:
            self.output.text = f"Gagal membaca file:\n{e}"


if __name__ == "__main__":
    RekapApp().run()
