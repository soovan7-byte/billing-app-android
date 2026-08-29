安卓打包说明

本项目通过 GitHub Actions 在线构建 Android APK，仓库内不包含任何本地构建脚本或环境配置。

仓库结构（与打包相关）
- main.py：安卓版主程序
- buildozer.spec：Buildozer 打包配置（arm64-v8a）
- .github/workflows/android.yml：GitHub Actions 构建工作流

GitHub 在线构建方式
1. 打开仓库的 Actions 页面；
2. 选择工作流 “Build Android APK”；
3. 点击 Run workflow，选择 main 分支；
4. 等待构建完成；
5. 在成功的 workflow run 中下载 android-apk Artifact；
6. 解压后得到 APK 安装到 Android 手机（需允许安装未知来源应用）。

发布版本
- 已发布的稳定版本见仓库 Releases 页面，附带的 APK 为经过 Android 真机验证的安装包；
- 在线工作流生成的为 debug APK，主要用于验证最新代码。

功能说明
- 数据本地存储在应用私有目录；
- 导出通过 Android SAF 系统文件选择器保存；
- 导入支持 json/csv/xlsx；
- 导入通过记录 ID 自动跳过重复记录。
