from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.utils import platform
from kivy.clock import Clock

from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import os

# 🔥 ANDROID PERMISSION
if platform == "android":
    from android.permissions import request_permissions, Permission


class RekapApp(App):

    def build(self):
        root = BoxLayout(orientation='vertical', padding=10, spacing=10)

        title = Label(
            text="Rekap Data Ruijie Pro",
            size_hint=(1, 0.1)
        )
        root.add_widget(title)

        self.result_label = Label(
            text="Menunggu file Excel...",
            size_hint_y=None,
            halign="left",
            valign="top"
        )
        self.result_label.bind(
            texture_size=lambda inst, val: setattr(inst, 'height', val[1])
        )

        scroll = ScrollView()
        scroll.add_widget(self.result_label)
        root.add_widget(scroll)

        btn = Button(text="Pilih File Excel", size_hint=(1, 0.1))
        btn.bind(on_press=self.buka_file)
        root.add_widget(btn)

        # 🔥 minta permission saat start
        Clock.schedule_once(self.request_android_permissions, 1)

        return root

    # ===============================
    # 🔐 REQUEST PERMISSION
    # ===============================
    def request_android_permissions(self, dt):
        if platform == "android":
            request_permissions([
                Permission.READ_EXTERNAL_STORAGE,
                Permission.WRITE_EXTERNAL_STORAGE
            ])

    # ===============================
    # 📂 buka file
    # ===============================
    def buka_file(self, instance):

        start_path = self.get_android_path()

        chooser = FileChooserListView(
            path=start_path,
            filters=["*.xlsx", "*.xls"],
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
    # 📱 path android
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

        rekap = defaultdict(int)

        for row in sheet.iter_rows(min_row=2, values_only=True):
            try:
                tanggal = str(row[11])[:10]
                harga = int(row[4])
                rekap[tanggal] += harga
            except:
                continue

        hasil = "===== HASIL REKAP =====\n\n"
        for tgl, total in sorted(rekap.items()):
            hasil += f"{tgl} : Rp {total:,}\n"

        self.result_label.text = hasil


if __name__ == "__main__":
    RekapApp().run()
