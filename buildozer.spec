[app]
title = 个人记账
package.name = personalaccounting
package.domain = org.openaiuser
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json,csv,xlsx,ttf,ttc,otf
version = 1.0
requirements = python3,kivy,openpyxl,et_xmlfile
orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True

[buildozer]
log_level = 2
warn_on_root = 1
