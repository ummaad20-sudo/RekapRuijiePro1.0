[app]

title = Rekap Ruijie Pro
package.name = rekapruijie
package.domain = com.ruijie
source.dir = .
source.include_exts = py,png,jpg,kv
version = 1.0
requirements = python3,kivy,openpyxl,et_xmlfile,plyer
orientation = portrait
fullscreen = 0

[buildozer]
log_level = 2

[app:android]
android.permissions = READ_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b
