[app]

title = Rekap Ruijie Pro
package.name = rekapruijie
package.domain = com.ruijie

source.dir = .
source.include_exts = py,png,jpg,kv

version = 1.0

requirements = python3,kivy,plyer,openpyxl,et_xmlfile

orientation = portrait
fullscreen = 0

# ===============================
# ANDROID CONFIG (STABLE)
# ===============================
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

android.api = 31
android.minapi = 21
android.sdk = 31
android.ndk = 25b

[buildozer]
log_level = 2
