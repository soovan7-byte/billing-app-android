[app]
title = 个人记账
package.name = personalaccounting
package.domain = org.openaiuser
source.dir = .
# 项目实际资源：main.py + simkai.ttf；如后续新增 kv/图片再补充扩展名
source.include_exts = py,ttf
version = 1.0
requirements = python3,kivy,openpyxl,et_xmlfile
orientation = portrait
fullscreen = 0
# 导入导出均使用 Android SAF 系统文件选择器，数据保存在应用私有目录，无需存储权限
# android.permissions =
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True

[buildozer]
log_level = 2
warn_on_root = 1
