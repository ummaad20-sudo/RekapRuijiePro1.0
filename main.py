import os

# =====================================
# ANDROID PERMISSIONS (PENTING)
# =====================================
try:
    from android.permissions import request_permissions, Permission
    request_permissions([
        Permission.READ_EXTERNAL_STORAGE,
        Permission.WRITE_EXTERNAL_STORAGE,
        Permission.MANAGE_EXTERNAL_STORAGE
    ])
except Exception:
    pass

# =====================================
# IMPORT
# =====================================
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

        title = Label(
            text="Rekap Excel",
            size_hint=(1, 0.1)
        )
        root.add_widget(title)

        btn = Button(
            text="Pilih File Excel",
            size_hint=(1, 0.1)
        )
        btn.bind(on_press=self.buka_file)
        root.add_widget(btn)

        # output scrollable
        self.output = TextInput(
            readonly=True,
            size_hint_y=None
        )
        self.output.bind(minimum_height=self.output.setter("height"))

        scroll = ScrollView()
        scroll.add_widget(self.output)
        root.add_widget(scroll)

        return root

    # =====================================
    # FILE PICKER
    # =====================================
    def buka_file(self, instance):
        try:
            filechooser.open_file(
                title="Pilih File Excel",
                path="/sdcard/Download",  # ⭐ penting Android
                filters=["*.xlsx"],       # ⭐ lebih kompatibel
                on_selection=self.file_selected
            )
        except Exception as e:
            self.output.text = f"Gagal membuka file picker:\n{e}"

    def file_selected(self, selection):
        if not selection:
            self.output.text = "Tidak ada file dipilih."
            return

        file_path = selection[0]
        self.proses_excel(file_path)

    # =====================================
    # PROSES EXCEL
    # =====================================
    def proses_excel(self, file_path):
        try:
            if not os.path.exists(file_path):
                self.output.text = f"File tidak ditemukan:\n{file_path}"
                return

            wb = load_workbook(file_path, data_only=True)
            sheet = wb.active

            rekap = defaultdict(int)

            for row in sheet.iter_rows(min_row=2, values_only=True):

                # guard kolom
                if not row or len(row) < 12:
                    continue

                tanggal = row[11]  # kolom L
                harga = row[4]     # kolom E

                if not tanggal or harga is None:
                    continue

                # konversi harga aman
                try:
                    nilai = int(float(str(harga).replace(",", "")))
                except Exception:
                    nilai = 0

                tgl_str = str(tanggal)[:10]
                rekap[tgl_str] += nilai

            # =====================================
            # OUTPUT
            # =====================================
            hasil = "=== REKAP PER TANGGAL ===\n\n"

            grand_total = 0
            for tgl in sorted(rekap.keys()):
                total = rekap[tgl]
                grand_total += total
                hasil += f"{tgl} : Rp {total:,}\n"

            hasil += "\n===========================\n"
            hasil += f"GRAND TOTAL : Rp {grand_total:,}\n"

            self.output.text = hasil

        except Exception as e:
            self.output.text = f"Gagal membaca file:\n{e}"


# =====================================
# RUN
# =====================================
if __name__ == "__main__":
    RekapApp().run()
