[app]

title = Rekap Voucher Ruijie
package.name = rekapruijie
package.domain = com.ruijie

source.dir = .
source.include_exts = py,png,jpg,kv

version = 1.0

requirements = python3,kivy,openpyxl,plyer

orientation = portrait
fullscreen = 0

android.permissions = READ_EXTERNAL_STORAGE

android.api = 34
android.minapi = 24
android.sdk = 31
android.ndk = 25b

[buildozer]
log_level = 2
