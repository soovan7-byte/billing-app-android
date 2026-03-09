安卓打包说明

这个项目已经整理成 Buildozer 可打包结构：
- main.py：安卓版主程序
- buildozer.spec：打包配置

建议打包方式一：GitHub Actions
1. 新建一个 GitHub 仓库。
2. 把整个 android_billing_app 文件夹里的内容上传到仓库根目录。
3. 打开 Actions，运行工作流“Build Android APK”。
4. 等待完成后，在 Artifacts 里下载 APK。

建议打包方式二：Ubuntu / WSL 本地打包
1. 安装 Ubuntu 或 WSL。
2. 安装 buildozer、Java、Android SDK/NDK 依赖。
3. 在项目目录运行：buildozer android debug
4. 生成的 APK 一般在 bin 目录。

功能说明
- 数据本地存储在应用私有目录。
- 导出默认保存到：
  - Windows：项目当前目录
  - Android：Download 目录（若权限允许）
- 导入支持 json/csv/xlsx。
- 导入会自动跳过重复记录。
- 查看记录和删除记录按最新时间自动排序。

说明
当前我无法在此环境直接编译出 APK，但项目已经整理成可打包版本。
