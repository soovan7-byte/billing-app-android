# 参与贡献指南

感谢你关注“无广告极简记账 App”。本项目是使用 Python 与 Kivy 开发的本地离线 Android 个人记账应用。提交 Issue 或 Pull Request 时，请尽量保持描述准确、范围清晰，并避免提交任何真实账单或隐私数据。

## Issue 提交要求

提交 Issue 前，请先确认：

- 问题与当前仓库代码或文档有关；
- 描述的是当前项目实际能力范围内的问题；
- 没有将云同步、账号系统、广告接入、自动同步银行卡、微信或支付宝自动记账等当前未实现能力作为已存在功能来描述；
- 不包含真实姓名、真实账单、银行卡信息、手机号、邮箱、地址、身份证件、密钥、签名文件或其他隐私数据；
- 如需展示数据格式，请使用脱敏后的示例数据。

建议 Issue 包含：

- 问题类型：Bug、文档、构建、导入导出、统计、Android 实机表现或其他；
- 简要标题；
- 清晰的复现步骤或改进建议；
- 期望结果；
- 实际结果；
- 相关截图或日志（如有，请先脱敏）。

## Bug 报告所需信息

报告 Bug 时，请尽量提供以下信息：

- 当前分支或提交号；
- 运行环境：桌面 Python 环境或 Android 真机；
- Android 设备型号和 Android 系统版本（如适用）；
- APK 来源：本地 Buildozer 构建或 GitHub Actions debug artifact；
- 复现步骤；
- 发生频率：必现、偶现或只出现过一次；
- 相关输入文件格式：JSON、CSV 或 XLSX（如涉及导入）；
- 错误日志、构建日志或应用提示信息；
- 是否涉及分类、统计、导入导出、删除记录或 Android SAF 文件选择器。

请不要上传真实账单文件。需要说明导入问题时，可构造最小化、脱敏后的示例文件。

## Pull Request 规范

提交 Pull Request 时，请遵守以下规则：

- 一个 PR 尽量只解决一个明确问题；
- 标题简洁说明变更范围，例如 `docs: update backup notes` 或 `fix: skip invalid import rows`；
- 在 PR 描述中说明修改动机、主要变更和已运行的检查；
- 不要在文档中宣传尚未实现的功能；
- 不要虚构发布日期、正式版本号、下载量、应用商店评分、Release 安装包或构建始终通过状态；
- 修改导入导出、Android SAF、数据结构、Buildozer 配置、权限、包名、版本号或 UI 行为时，应明确说明影响范围；
- 不要提交生成的 APK、AAB、签名文件、Buildozer 构建缓存或用户账单数据。

## 修改 `main.py` 后的基础检查命令

如果修改了 `main.py`，至少运行以下基础检查：

```bash
python -m py_compile main.py
```

提交前建议同时运行：

```bash
git diff --check
git status --short
git diff --name-only
```

这些检查不能替代 Android 真机测试。涉及 Android 行为、文件选择、导入导出或触控界面的改动，应进行 Android 真机验证。

## Android 实机测试说明

当前项目面向 Android 手机使用。以下变更建议进行实机测试：

- 记账、明细、统计或设置页面交互；
- 底部导航；
- 分类管理；
- 单条记录删除和清空记录；
- 月度统计、年度统计、分类饼图和分类占比；
- JSON、CSV、XLSX 导入；
- Excel、CSV 和完整 JSON 备份导出；
- Android SAF 系统文件选择器；
- Buildozer 配置、Android 权限、包名、版本号或依赖变更。

实机测试记录建议包含：

- 设备型号；
- Android 版本；
- APK 构建方式；
- 测试的主要流程；
- 是否覆盖导入导出和备份恢复。

## 隐私数据与敏感文件禁止提交

请不要提交以下内容：

- `records.json`；
- `categories.json`；
- 真实账单导出的 Excel、CSV 或 JSON 文件；
- 完整 JSON 备份文件；
- 导入导出临时文件；
- Android keystore、`.jks`、`.keystore` 或签名配置；
- API Key、Token、密码、私钥、证书或其他密钥；
- 包含个人身份信息的截图、日志或测试数据。

如果不确定某个文件是否可以提交，请先不要提交，并在 Issue 或 PR 中说明用途。
