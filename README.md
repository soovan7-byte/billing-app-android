# 无广告极简记账 App

一个使用 Python 与 Kivy 开发的本地离线 Android 记账应用，无广告、无需账号，专注于简单记录和查看个人消费。

**Minimal Ad-Free Expense Tracker**

![Python](https://img.shields.io/badge/Python-3.11-blue)
![Kivy](https://img.shields.io/badge/Kivy-App%20Framework-4B8BBE)
![Android](https://img.shields.io/badge/Android-debug%20APK-3DDC84)
![GitHub Actions](https://img.shields.io/badge/GitHub%20Actions-manual%20build-2088FF)
![Ad-Free](https://img.shields.io/badge/Ad--Free-current%20code-brightgreen)
![Offline](https://img.shields.io/badge/Offline-local%20first-orange)

## 项目简介

**无广告极简记账 App** 是一个面向个人日常消费记录的本地离线 Android 记账应用。项目使用 Python 和 Kivy 开发，数据优先保存在应用本地目录，适合用于快速记录金额、分类、备注和日期，并查看明细与统计结果。

项目当前代码具备以下特点：

- 无广告：当前项目代码不包含广告功能或第三方广告 SDK；
- 本地离线运行：不依赖远程服务器保存账单；
- 无需注册账号；
- 无需登录；
- 不包含云同步；
- 用户自行管理、导出和备份自己的数据；
- 卸载应用前应先导出完整 JSON 备份。

> 说明：这里的“无广告”仅表示当前项目代码不包含广告功能或广告 SDK，不代表对未来版本、第三方修改版或重新打包版本作出保证。

## 设计目标

本项目强调以下设计方向：

- 无广告；
- 极简；
- 本地离线；
- 快速记账；
- 数据由用户掌控；
- 不加入不必要的账号、社交和推荐功能。

项目目标是保留个人记账所需的核心能力，避免将简单记账流程复杂化。

## 主要功能

- 快速记录金额、分类、备注和日期；
- 自定义消费分类；
- 本月总支出和记录数量；
- 最近账单明细；
- 查看完整记录详情；
- 单条记录安全删除；
- 月度统计；
- 年度统计；
- 分类饼图；
- 分类金额及占比；
- Excel 账单表格导出；
- CSV 账单表格导出；
- JSON 账单和分类完整备份；
- JSON、CSV、XLSX 数据导入；
- 重复记录自动跳过；
- Android SAF 系统文件选择器；
- 本地离线数据存储；
- 无广告、无账号、无云服务依赖。

## 应用页面

### 记账

记账页用于快速新增一条消费记录：

- 输入金额；
- 选择分类；
- 填写消费备注；
- 选择日期；
- 使用“记一笔”保存；
- 保存成功后显示非阻断反馈。

### 明细

明细页用于查看和管理最近账单：

- 查看最近记录；
- 点击记录查看完整详情；
- 删除错误记录；
- 删除前二次确认。

### 统计

统计页用于查看消费汇总和分类占比：

- 月度统计；
- 年度统计；
- 总支出；
- 有效记录数量；
- 分类饼图；
- 分类金额及占比；
- 年度 1 月至 12 月支出明细。

### 设置

设置页用于管理分类、导入导出和危险操作：

- 分类管理；
- 数据导入；
- Excel、CSV、JSON 导出；
- 完整备份；
- 数据数量概览；
- 清空所有账单。

## 应用截图

截图将在后续版本中补充。

## 数据与隐私

- 账单和分类保存在应用本地；
- 项目本身不会主动上传账单数据；
- 不包含用户账号系统；
- 不包含广告 SDK；
- 不包含云同步；
- 不依赖远程服务器保存账单；
- Android 卸载应用通常会删除应用私有目录中的数据；
- 卸载前应导出完整 JSON 备份；
- 导出后的文件由用户自行保存和保护；
- 不要把包含真实账单的文件提交到公开 GitHub 仓库。

## 数据导入与导出

### Excel 导出

Excel 导出用于查看和整理账单表格：

- 文件格式为 `.xlsx`；
- 只包含账单记录；
- 不包含自定义分类配置。

### CSV 导出

CSV 导出用于表格软件或其他程序读取：

- 文件格式为 `.csv`；
- 只包含账单记录；
- 不包含自定义分类配置。

### 完整 JSON 备份

完整 JSON 备份用于迁移和恢复应用数据：

- 包含账单 `records`；
- 包含分类 `categories`；
- 适合在换机、重装或卸载应用前保存完整数据。

完整备份 JSON 简化示例：

```json
{
  "records": [
    {
      "姓名/备注": "午餐",
      "分类": "饮食正餐",
      "金额": 25.0,
      "日期": "2026-07-19",
      "记录时间": "2026-07-19 12:00:00"
    }
  ],
  "categories": [
    "饮食正餐",
    "交通"
  ]
}
```

### 导入兼容格式

当前导入功能兼容以下数据来源：

- 旧账单列表 JSON；
- 独立分类列表 JSON；
- 包含 `records` 和 `categories` 的完整备份 JSON；
- CSV；
- XLSX。

导入时会进行基础校验：

- 导入会校验记录格式；
- 无效记录会被跳过；
- 重复记录会自动跳过；
- 导入不会保证恢复其他版本自行增加的未知字段。

## 使用说明

### 安装与运行

当前仓库根目录只维护一个 `README.md` 作为项目首页，主要提供源码和 GitHub Actions 手动构建 debug APK 的能力。项目未声明已经发布到应用商店，也未在 GitHub Release 中提供正式安装包。

如果通过 GitHub Actions 构建 APK，可在仓库页面进入 **Actions**，选择 Android APK 构建工作流，手动触发构建，然后在构建产物中下载 debug APK。

### 日常记账流程

1. 打开应用后进入“记账”页；
2. 输入金额；
3. 选择或先在设置中新增消费分类；
4. 填写备注；
5. 选择日期；
6. 点击“记一笔”保存；
7. 在“明细”页查看最近记录；
8. 在“统计”页查看月度、年度和分类统计。

### 备份建议

- 定期从“设置”页导出完整 JSON 备份；
- 更换手机、清除应用数据或卸载应用前，先导出完整 JSON 备份；
- Excel 和 CSV 适合查看账单表格，但不包含分类配置；
- 如果需要迁移和恢复，请优先使用完整 JSON 备份。

## 开发者说明

### 技术栈

- Python；
- Kivy；
- Buildozer；
- openpyxl；
- Android SAF 系统文件选择器；
- GitHub Actions 手动构建 debug APK。

### 项目结构

```text
.
├── .github/workflows/android.yml   # GitHub Actions 手动构建 Android debug APK
├── buildozer.spec                  # Buildozer Android 打包配置
├── main.py                         # Kivy 应用主代码
├── simkai.ttf                      # 应用使用的本地字体文件
└── README_安卓打包说明.txt          # 早期 Android 打包说明
```

### 本地依赖

应用源码依赖 Python、Kivy 和 openpyxl。Android 打包依赖 Buildozer、Android SDK、Java 以及相关系统编译工具。

可参考 GitHub Actions 工作流中的环境配置来准备构建环境。

### 本地运行参考

在已安装 Python 依赖的桌面环境中，可尝试直接运行：

```bash
python main.py
```

桌面运行主要用于开发调试；实际 Android 行为仍应以真机测试为准。

### Android debug APK 构建参考

在具备 Buildozer 环境的机器上，可执行：

```bash
buildozer android debug
```

构建产物通常位于 `bin/` 目录。不同系统的 Android SDK、NDK、Java、Buildozer 缓存状态可能影响构建结果。

## GitHub Actions 手动构建

仓库包含一个 GitHub Actions 工作流，用于手动构建 Android debug APK：

- 工作流触发方式为 `workflow_dispatch`；
- 使用 Ubuntu 22.04；
- 设置 Python 3.11；
- 设置 Java 17；
- 安装 Android SDK；
- 安装 Buildozer 和兼容版本 Cython；
- 执行 `buildozer android debug`；
- 构建成功后上传 `bin/*.apk` 作为 workflow artifact。

该工作流用于生成 debug APK，不等同于正式发布版本。

## 维护说明

### 文档维护原则

- README 中的功能描述应与当前代码实际能力保持一致；
- 不要写入尚未实现的功能；
- 不要承诺应用商店发布、Release 安装包、云同步或第三方账号能力；
- 不要把真实账单、真实导出文件或个人隐私数据提交到仓库；
- 如果后续新增功能，应同步更新 README 中的功能、隐私和导入导出说明。

### 代码维护提醒

本次文档整理不修改应用业务代码、UI、数据结构、Android SAF、构建逻辑、包名、安装标题或版本号。后续维护时如需修改这些内容，建议单独提交并在提交说明中明确变更范围。

### 实机测试状态

当前 `new-something` 分支版本已经完成 Android 手机实机测试。后续如修改 Android 权限、文件导入导出、Buildozer 配置或 Kivy UI，建议重新进行真机验证。

## 当前未包含的能力

为避免误解，当前项目不包含以下能力：

- 云同步；
- 银行卡自动同步；
- 微信或支付宝自动记账；
- 多人共享账本；
- 预算提醒；
- 账号系统；
- AI 自动分类；
- 应用商店正式发布；
- GitHub Release 正式安装包。

## 适用人群

- 希望本地记录个人消费的用户；
- 不需要账号系统和云同步的用户；
- 希望查看月度、年度和分类统计的用户；
- 需要导出表格或完整备份自己账单数据的用户；
- 想了解 Python + Kivy Android 应用结构的开发者。
