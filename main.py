# -*- coding: utf-8 -*-
# Rekap Ruijie Pro — Versi Pydroid3 Review
# Pembuat: Jun

import os
from threading import Thread

from kivy.app import App
from kivy.lang import Builder
from kivy.clock import Clock
from kivy.properties import StringProperty, NumericProperty

from openpyxl import load_workbook

KV = '''
<Root>:
    orientation: 'vertical'
    padding: dp(16)
    spacing: dp(12)

    Label:
        text: 'Rekap Ruijie Pro (by Jun)'
        font_size: '20sp'
        size_hint_y: None
        height: dp(40)

    Button:
        text: 'Pilih File Excel'
        size_hint_y: None
        height: dp(48)
        on_release: root.pick_file()

    Label:
        text: root.file_label
        halign: 'left'
        valign: 'middle'
        text_size: self.size

    Label:
        text: 'Total Baris: ' + str(root.total_rows)

    Label:
        text: 'Total Rupiah: ' + root.total_rupiah
'''


def format_rupiah(value):
    try:
        return 'Rp {:,}'.format(int(value)).replace(',', '.')
    except Exception:
        return 'Rp 0'


class Root(App.get_running_app().__class__ if App.get_running_app() else object):
    pass


from kivy.uix.boxlayout import BoxLayout


class Root(BoxLayout):
    file_label = StringProperty('Belum ada file dipilih')
    total_rows = NumericProperty(0)
    total_rupiah = StringProperty('Rp 0')

    def pick_file(self):
        # Untuk Pydroid3: hardcode path contoh
        # Silakan ganti path ini sesuai lokasi file Excel Anda
        example_path = '/sdcard/Download/Voucher.xlsx'

        if not os.path.exists(example_path):
            self.file_label = 'File tidak ditemukan di: ' + example_path
            return

        self.file_label = 'Memproses: ' + os.path.basename(example_path)
        Thread(target=self.process_file, args=(example_path,), daemon=True).start()

    def process_file(self, path):
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active

            total = 0
            rows = 0

            data_iter = list(ws.iter_rows(min_row=2, values_only=True))

            for row in data_iter:
                try:
                    price = row[4] if len(row) > 4 and row[4] is not None else 0
                    if isinstance(price, (int, float)):
                        total += price
                        rows += 1
                except Exception:
                    pass

            Clock.schedule_once(lambda dt: self.finish(rows, total))

        except Exception as e:
            Clock.schedule_once(lambda dt: setattr(self, 'file_label', f'Error: {e}'))

    def finish(self, rows, total):
        self.total_rows = rows
        self.total_rupiah = format_rupiah(total)
        self.file_label = 'Selesai diproses'


class RekapApp(App):
    def build(self):
        Builder.load_string(KV)
        return Root()


if __name__ == '__main__':
    RekapApp().run()
