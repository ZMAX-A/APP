# 颜佳AI Android 平板自动化

这是一个针对已安装应用的黑盒 Appium 3 + UiAutomator2 + pytest 项目，不需要颜佳AI源码。

当前实机基线：

- 包名：`com.xiaofutech.yanjia_ai`
- 启动 Activity：`.activity.SplashActivity`
- 已验证版本：`2.1.9(Build 8)`（versionCode `20109`）
- 已验证设备：Android 16 / API 36，小米平板，横屏 `3200 x 2136`
- Appium：`3.5.2`
- UiAutomator2 driver：`7.6.2`
- Python client：`5.3.1`

## 已自动化的只读路径

- 登录或恢复已有会话，选择门店并进入首页
- 首页四个入口的页面契约
- 顾客档案列表及首张顾客卡片
- 首位顾客详情页
- 案例库页面
- 设置与个人中心字段
- 已有影像进入阅览页（扩展套件，依赖种子数据）

项目现在使用 Excel 数据驱动执行。pytest 只有一个通用测试入口，用例、步骤、定位器、输入和断言全部从 `test_case.xlsx` 读取；执行结果会写回 Excel，并同时生成 Allure 报告。

## Excel 工作表

- `自动化测试用例`：一行一个用例。`是否执行=是` 才会运行；主表保存最新结果、耗时、错误、运行编号和 Allure 报告路径。
- `自动化执行步骤`：一行一个步骤，通过 `用例ID` 关联主表。步骤按 `步骤序号` 执行。
- `自动化关键字说明`：可以使用的操作、断言、定位器和变量写法。
- `执行记录`：每次执行都追加一条历史记录，不覆盖过去的运行结果。

当前 13 个已有真实定位器的用例已迁移到步骤表。以后在现有关键字范围内新增、复制或修改用例，不需要修改 Python 测试脚本。

## 首次准备

1. 平板开启开发者选项和 USB 调试，并授权当前电脑。
2. 小米/HyperOS 设备还需要开启“USB 调试（安全设置）”以及“通过 USB 安装”。
3. 保持平板解锁，首次创建 Appium Session 时允许安装以下官方辅助包：
   - `io.appium.settings`
   - `io.appium.uiautomator2.server`
   - `io.appium.uiautomator2.server.test`
4. 执行环境检查：

```powershell
.\scripts\doctor.ps1
```

当前 `env.txt` 支持纯值和带标签两种两行格式，例如 `账号：xxx`、`密码：xxx`。建议最终迁移为本地 `.env`：

```dotenv
YANJIA_USERNAME=测试账号
YANJIA_PASSWORD=测试密码
YANJIA_STORE_NAME=
ANDROID_UDID=
```

门店选择规则：`YANJIA_STORE_NAME` 填了门店名就直接使用该门店；留空时，运行时会在终端列出 App 返回的所有门店，输入序号选择（非交互环境如 CI 中自动选择第一个）。

`env.txt`、`.env`、报告、截图、APK 均已加入 `.gitignore`。程序缺少凭据时只报告缺失键，不回显值。

## 安装 Python 依赖

已有依赖时可直接运行。新机器执行：

```powershell
python -m pip install -e ".[dev]"
```

也可以安装 `uv` 后执行 `uv sync --extra dev`。

## 运行 Excel 用例

一般无需手动启动 Appium；运行入口会按需自动启动并在结束后清理。需要单独调试服务时可执行：

```powershell
.\scripts\start-appium.ps1
```

部分小米/HyperOS 设备在辅助包已安装后，仍会拦截 Appium Settings 服务的重复初始化。确认以下三个官方辅助包已成功安装后，可以在本地 `.env` 中启用：

```dotenv
APPIUM_SKIP_DEVICE_INITIALIZATION=true
APPIUM_SKIP_SERVER_INSTALLATION=true
```

这两个开关默认关闭；新设备首次运行时应保持关闭，以便 Appium 完成设备初始化和版本检查。

运行 Excel 中所有 `是否执行=是` 的用例，并将结果直接写回原工作簿：

```powershell
.\scripts\run-excel.ps1
```

这个入口会依次完成 Excel 离线校验、ADB 设备与安装包预检、Appium 就绪检查、pytest 执行、Excel 回写和 Allure 生成。已有 Appium 会直接复用；自动启动的 Appium 只由本次脚本清理。

只校验 Excel、不连接平板：

```powershell
.\scripts\validate-excel.ps1
```

调试时需要保留脚本自动启动的 Appium：

```powershell
.\scripts\run-excel.ps1 -KeepAppium
```

当前平板已有受控顾客和影像数据时：

```powershell
.\scripts\run-excel.ps1 -RunSeeded
```

只运行某条或某组用例：

```powershell
.\scripts\run-excel.ps1 -CaseId "TC-HOME-001"
.\scripts\run-excel.ps1 -CaseId "TC-HOME-*"
.\scripts\run-excel.ps1 -Tags "smoke,readonly" -RunSeeded
```

如果使用更新后的单表 Web/Playwright 用例文件，当前仅支持已迁移的 Android 登录用例：

```powershell
.\scripts\run-excel.ps1 -CaseId "TC-LOGIN-001,TC-LOGIN-002,TC-LOGIN-003" -NoWriteBack
```

单表中的 CSS 定位器和 URL 断言不会直接用于原生 Android；执行器会将已支持的登录用例映射为真实 resource-id 和 Activity。未选择用例时会拒绝执行，以免误把 Web 用例当成 Android 用例。

运行后立即打开报告：

```powershell
.\scripts\run-excel.ps1 -RunSeeded -OpenReport
```

稍后打开最近一次报告：

```powershell
.\scripts\open-allure.ps1
```

需要保留主工作簿、不原地回写时，可以指定结果副本：

```powershell
.\scripts\run-excel.ps1 -ExcelOutput "reports\本次测试结果.xlsx"
```

运行前必须关闭正在打开 `test_case.xlsx` 的 WPS/Excel 窗口，否则脚本会拒绝执行，避免工作簿损坏。每次写入前都会在 `reports/excel-backups` 自动备份。

## 安全门禁

- `requires_seed`：默认跳过，需要 `--run-seeded`。
- `mutating`：默认跳过，需要 `--allow-mutation`。
- `destructive`：默认跳过，必须同时提供 `--allow-mutation --allow-destructive`。
- 当前没有启用任何删除步骤；破坏性用例仍保持 `是否执行=否`。
- 失败时默认只保存测试节点和 Activity 等非敏感元数据。
- 顾客页、个人中心、截图和页面树可能包含个人信息。只有明确设置
  `YANJIA_CAPTURE_SENSITIVE_ARTIFACTS=true` 才会保存截图与 XML。

不要在生产租户上开启写入或删除门禁。

## 修改或新增用例

1. 在 `自动化测试用例` 新增或复制一行，设置唯一 `用例ID`、标签和 `是否执行=是`。
2. 在 `自动化执行步骤` 为这个 `用例ID` 添加步骤，填写步骤序号、操作类型、定位器、输入或断言。
3. 在 `自动化关键字说明` 查看允许的格式。操作和断言列已提供下拉选项，定位器优先使用 `id=完整resource-id`。
4. 需要兼容不同版本控件时，可按优先级填写候选定位器，例如 `id=com.example:id/save || accessibility_id=保存`。
5. 输入和期望值可以引用 `.env` 变量，例如 `${SEEDED_CASE_TAG}`。
6. 保存并关闭 Excel，然后执行 `run-excel.ps1`。

只有在需要全新的底层交互能力、且“自动化关键字说明”中没有对应关键字时，才需要扩展执行引擎。

重新补齐工作簿结构或恢复当前已实现步骤时：

```powershell
.\scripts\prepare-excel.ps1
.\scripts\prepare-excel.ps1 -ReplaceCurrentSteps
```

第二条命令会替换当前 13 个内置基线用例的步骤；正常日常执行不需要运行它。

## 本地质量检查

```powershell
.\.venv\Scripts\python.exe -m pyright
.\.venv\Scripts\python.exe -m ruff check src tests scripts
.\.venv\Scripts\python.exe -m pytest tests\unit -q
```

## 已知边界

- 影像阅览首次进入可能需要下载较大资源，应放在扩展套件，不阻塞最短 Smoke。
- 搜索、筛选、编辑和删除依赖受控测试数据；未建立数据创建/恢复能力前不会默认执行。
- 颜佳AI当前页面均能从原生 accessibility 树取得稳定 `resource-id`，不使用 CSS 或 URL 断言。
- Appium 仅监听 `127.0.0.1`，不要启用全局 `--relaxed-security`。
- 只有带 `readonly` 标签的用例遇到 Appium 会话/连接故障时才自动重试一次；业务断言、写入和删除用例不会自动重试。
- Allure 报告保留历史趋势，并记录平板型号、Android/SDK 和颜佳AI版本；失败会按业务断言、定位等待、基础设施和配置分类。
