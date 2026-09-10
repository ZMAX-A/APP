# 颜佳AI Android 平板自动化

这是一个针对已安装应用的黑盒 Appium 3 + UiAutomator2 + pytest 项目，不需要颜佳AI源码。

当前自动化项目正式版本：`1.0.0`。

当前实机基线：

- 包名：`com.xiaofutech.yanjia_ai`
- 启动 Activity：`.activity.SplashActivity`
- 已验证版本：`2.2.0(Build 2)`（versionCode `20200`）
- 已验证设备：Android 16 / API 36，小米平板，横屏 `3200 x 2136`
- Appium：`3.5.2`
- UiAutomator2 driver：`7.6.2`
- Python client：`5.3.1`

V1.0.0 发布门禁：2026-09-10 在上述设备和应用版本完成 Run `20260910_111128`，84/84 条启用用例全部通过；发布前另有 204/204 条单元测试、Ruff、Pyright 和 PowerShell 语法检查通过。仓库中的工作簿作为可公开复用的测试基线，实际结果统一重置为 `NOT_RUN`，不包含本机运行历史、绝对报告路径、账号密码或完整顾客手机号。

## 已自动化的只读路径

- 登录或恢复已有会话，选择门店并进入首页
- 首页四个入口的页面契约
- 顾客档案列表及首张顾客卡片
- 首位顾客详情页
- 影像备注与咨询单备注的空格输入校验（真实点击提交，验证备注记录数不增加）
- 影像阅览页查看案例、查看报告及不保存返回顾客详情（只读种子数据）
- 案例库页面
- 设置与个人中心字段
- 已有影像进入阅览页（扩展套件，依赖种子数据）

项目现在使用 Excel 数据驱动执行。pytest 只有一个通用测试入口，用例、步骤、定位器、输入和断言全部从 `test_case.xlsx` 读取；执行结果会写回 Excel，并同时生成 Allure 报告。

## Excel 工作表

- `自动化测试用例`：一行一个用例。`是否执行=是` 才会运行；主表保存最新结果、耗时、错误、运行编号和 Allure 报告路径。
- `自动化执行步骤`：一行一个步骤，通过 `用例ID` 关联主表。步骤按 `步骤序号` 执行。
- `自动化关键字说明`：可以使用的操作、断言、定位器和变量写法。
- `执行记录`：每次执行都追加一条历史记录，不覆盖过去的运行结果。

当前标准步骤表已登记 85 个用例、909 个步骤（84 条启用用例共 903 个步骤，另有未启用的 `TC-IMAGE-001` 6 个步骤）；工作簿 84 条启用用例已全部支持，P0/P1/P2/P3 覆盖分别为 15/15、62/62、6/6、1/1。以后在现有关键字范围内新增、复制或修改用例，不需要修改 Python 测试脚本。

当前工作簿的 `自动化执行步骤` 与代码目录保持同步；兼容加载器仍可支持未写入标准步骤表的旧版单表 Android 用例，`BACKLOG` 用例不会被误当成可执行 Android 用例。对于完全不含步骤表的旧版单表 Web/Playwright 工作簿，仍必须显式选择已迁移的用例 ID。

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
APPIUM_STARTUP_TIMEOUT_SECONDS=60
# 仅供隔离顾客写入测试；值可能包含个人信息，不要提交
YANJIA_MUTATION_CUSTOMER_QUERY=
# 精确手机号只读搜索可以复用同一个受控顾客，不复制明文
SEEDED_CUSTOMER_PHONE=${YANJIA_MUTATION_CUSTOMER_QUERY}
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

一般无需手动启动 Appium；运行入口会按需自动启动并在结束后清理。自动启动默认等待 60 秒，可通过 `.env` 的 `APPIUM_STARTUP_TIMEOUT_SECONDS` 或 `-AppiumStartupTimeoutSeconds` 在 5~300 秒范围内覆盖。需要单独调试服务时可执行：

```powershell
.\scripts\start-appium.ps1
```

部分小米/HyperOS 设备在辅助包已安装后，仍会拦截 Appium Settings 服务的重复初始化。确认以下三个官方辅助包已成功安装后，可以在本地 `.env` 中启用：

```dotenv
APPIUM_SKIP_DEVICE_INITIALIZATION=true
APPIUM_SKIP_SERVER_INSTALLATION=true
```

这两个开关默认关闭；新设备首次运行时应保持关闭，以便 Appium 完成设备初始化和版本检查。只读用例遇到可恢复的 UiAutomator2 传输或会话错误时，自动重试会忽略 `APPIUM_SKIP_SERVER_INSTALLATION` 一次，重新启用设备端 server 版本检查与必要安装后创建新会话；写入和删除用例不会执行这类即时重试；全量启动器的第二轮复跑另按下文规则执行。

自动登录后，`restart_to_home` 会兼容标准 `android:id/alertTitle` 与 OEM `*:id/alertTitle`（例如 MIUI 密码管理器），但仍仅在标题同时包含“账号”“密码”和“保存/存储”语义时点击 `android:id/button2` 取消保存；其他系统弹窗保持原状，等待人工判断。

运行 Excel 中 `是否执行=是` 且已经接入 Android 执行源的用例，并将结果直接写回原工作簿：

```powershell
.\scripts\run-excel.ps1
```

这个入口会依次完成 Excel 离线校验、ADB 设备与安装包预检、Appium 就绪检查、pytest 执行、Excel 回写和 Allure 生成。已有 Appium 会直接复用；自动启动的 Appium 只由本次脚本清理。启动超时时同时保留 `reports/appium/<RunId>.stdout.log` 和 `.stderr.log` 路径，便于区分服务加载慢、入口退出和驱动错误。

未安装 Allure CLI 时，入口仍会执行 pytest 并保留 `reports/allure-results` 原始结果，只跳过 HTML 报告生成；使用 `-OpenReport` 仍要求本机已安装 Allure CLI。

只校验 Excel、不连接平板：

```powershell
.\scripts\validate-excel.ps1
```

查看 Android 用例目录和当前工作簿覆盖率（不连接平板、不写回 Excel）：

```powershell
.\.venv\Scripts\python.exe scripts\report_android_coverage.py --source single_sheet
# canonical「自动化执行步骤」目录
.\.venv\Scripts\python.exe scripts\report_android_coverage.py --source step_sheet
```

报告会区分已接入、尚未迁移和未登记用例；发布门禁可增加
`--fail-on-p0-backlog`。当前目录由 `src/yanjia_automation/excel/android_catalog.py` 统一维护，单表加载器、`prepare_excel.py` 和 `normalize_cases.py` 会在启动时校验目录一致性。
当前基线为：`single_sheet` 支持 76/84 个启用用例（P1 为 56/62、P2 为 4/6、P3 为 1/1），`step_sheet` 支持 84/84（P1 为 62/62、P2 为 6/6、P3 为 1/1）；两者 P0 均为 15/15。标准步骤表可离线装载 84 条启用用例、903 个启用步骤；步骤表物理总量为 909 条，其中 6 条属于当前未启用的 `TC-IMAGE-001`。其中 24 条顾客详情写入或自包含删除校验还要求本地专用顾客查询和本次只读预检标记。

2026-08-28 完成 P2 `TC-CUSTOMER-007` 的单表/标准步骤表双执行源迁移，使用只读日期范围 `2015-02-04~2015-02-13` 并严格校验日期文本和空状态。Android 16/API 36、应用 2.1.9(Build 8) 真机连续两次在确认该空区间后跳转到 `.LoginActivity`：Run ID `20260828_104057` 和 `20260828_105211` 均在日期文本断言前失败；同设备、同版本的正向对照 `TC-CUSTOMER-008` Run ID `20260828_110414` 为 `1 passed in 334.30s`。因此当前证据指向空日期区间触发的产品自动登出缺陷，自动化仍保留严格失败，不把登出视为通过；三次运行均使用 `-NoWriteBack` 且未修改业务数据。

同日完成 P2 `TC-HOME-006` 的单表/标准步骤表双执行源迁移：在首页输入手机号片段 `186`，进入 `.CustomerRecordsActivity` 后先校验至少一张顾客卡片，再校验首个可见 `a_records_unique_tv` 包含该片段；不采集或记录完整手机号。Android 16/API 36、应用 2.1.9(Build 8) 真机 Run ID `20260828_143236` 为 `1 passed in 245.79s`，使用 `-RunSeeded -NoWriteBack`，运行前后工作簿 SHA-256 一致且未修改业务数据。

同日完成 P1 `TC-CUSTOMER-022` 的单表/标准步骤表双执行源迁移：打开其他筛选后选择男和 18-25 岁、输入临时备注 `123`，先逐项校验条件已设置，再点击重置并严格校验两项 `selected=false`、备注框 `text=搜索备注记录`、筛选面板仍可见；全程不点击“确定”或提交筛选。首轮 Run ID `20260828_151455` 因沿用错误占位文本“请输入备注”在第 15 步严格失败；按真机 `text/hint` 精确值修正后，Android 16/API 36、应用 2.1.9(Build 8) Run ID `20260828_154155` 为 `1 passed in 265.06s`。两轮均使用 `-RunSeeded -NoWriteBack`，工作簿哈希未变化且未修改业务数据。

同日完成 P3 `TC-HOME-007` 的单表/标准步骤表双执行源实现：完整手机号只通过 `${SEEDED_CUSTOMER_PHONE}` 在本地运行时解析，进入 `.CustomerRecordsActivity` 后严格校验顾客卡片数等于 1，再对 `a_records_unique_tv` 执行 `text_equals`；工作簿、日志和报告均不保存手机号明文。离线目标校验为 1 条用例、6 个步骤、0 错误/0 警告，全量单元测试 149 条通过，`step_sheet` 启用 backlog 清零。2026-08-31 使用本地受控顾客完成最终真机验收：Run `20260831_134545` 首次会话遇到 UiAutomator2 `socket hang up`，自动重试后因瞬时回到 `.LoginActivity` 严格失败；同环境正向对照 `TC-HOME-006` Run `20260831_135755` 为 `1 passed in 250.94s`，刷新登录状态后的确认 Run `20260831_140713` 为 `1 passed in 259.33s`。三轮均使用 `-RunSeeded -NoWriteBack`，工作簿 SHA-256 前后保持 `8CE3A736F754FE7F83BF6491D646235D67EBE73573BF14CDB6B8041075F20634`，未修改业务数据；当前状态为“实现完成、真机验收通过”。

2026-08-31 启动 APP5 稳定性阶段：Appium 启动等待默认调整为可配置的 60 秒；只读用例发生可恢复基础设施异常时，重试会保留本次命令行安全门禁，重新启用 UiAutomator2 server 安装检查并创建新会话，写入和破坏性用例仍不自动重试。离线验证结果为全量单元测试 `152 passed`、Ruff 通过、Pyright 0 错误、PowerShell 脚本解析通过，标准步骤表启用覆盖保持 84/84、backlog 为 0。2026-09-01 在 Android 16/API 36、应用 2.1.9(Build 8) 上执行只读确认 Run `20260901_091615`：新 Appium 进程在 60 秒配置窗口内就绪，`TC-HOME-007` 为 `1 passed in 371.57s`，Allure 状态为 passed，Appium stderr 为空且运行后进程已清理；运行前后工作簿 SHA-256 均为 `AFE77B41ACB4422D7A184CED48D8AB7A4F628436EDB52237B17CA873D6DDCB12`。本轮使用 `-RunSeeded -NoWriteBack`，未修改 Excel 或业务数据。

同日完成 APP5 可控故障恢复真机验证：最终 Run `20260901_095447` 在本轮 UiAutomator2 server 存活、已观察 `.MainActivity` 且随后进入 `.CustomerRecordsActivity` 后，仅停止一次 `io.appium.uiautomator2.server.test` 与 `io.appium.uiautomator2.server`。Allure 异常附件匹配 `socket hang up` 和 `could not proxy command`，`infrastructure_retry=1` 且自动重试附件为 1 个，Appium 日志确认 instrumentation 进程已停止；恢复分支重新创建会话后，`TC-HOME-007` 最终为 `1 passed in 239.10s`。Appium stderr 为空、运行后进程已清理，工作簿 SHA-256 前后保持 `AFE77B41ACB4422D7A184CED48D8AB7A4F628436EDB52237B17CA873D6DDCB12`，未修改 Excel 或业务数据。此前 Run `20260901_093304` 和 `20260901_094554` 均为通过但未触发重试的监控校准，不作为故障恢复证据。

同日完成 APP5 连续只读冒烟回归：`TC-HOME-006/007` 在 3 次独立 Appium 冷启动 Run `20260901_101104`、`20260901_101858`、`20260901_102714` 中均为 2/2 passed，合计 6/6 passed；三轮 pytest 耗时分别为 234.91、262.01、244.17 秒，中位数 244.17 秒。Allure 中自动重试为 0，未检测到 `.LoginActivity`，三轮 Appium stderr 均为空；工作簿 SHA-256 前后保持 `AFE77B41ACB4422D7A184CED48D8AB7A4F628436EDB52237B17CA873D6DDCB12`，报告与日志未检出种子手机号明文。该批次未发生自然基础设施异常，因此不计算自动恢复成功率；自动恢复结论仍由前述可控故障 Run 独立支撑。

2026-09-02 完成 APP5 全量只读真机基线：首次 Run `20260902_091127` 在 5 个 UiAutomator2 传输错误后出现 53 个凭据级联失败，仅 1 passed、25 skipped，因此不作为功能基线；随后 `TC-LOGIN-007` 控制 Run `20260902_094558` 为 1 passed，确认正向登录恢复。有效重试 Run `20260902_095446` 收集 84 条，实际执行 59 条只读用例并得到 56 passed、3 failed、25 skipped in 1099.56s；23 条写入和 2 条删除用例均由安全门禁跳过。3 个失败为 `TC-CUSTOMER-007` 日期文本未符合期望，以及 `TC-DETAIL-023/036` 空格输入后提交按钮仍启用；本轮无基础设施错误、自动重试标签为 0。工作簿 SHA-256 前后保持 `AFE77B41ACB4422D7A184CED48D8AB7A4F628436EDB52237B17CA873D6DDCB12`，三次 Run 的报告和日志未检出种子手机号明文。

2026-08-27 已在 Android 16/API 36 平板完成 `TC-DETAIL-020/037` 标准步骤表正式回归：Run ID `20260827_094813`，`2 passed in 228.88s`，并在独立只读会话中确认影像标签恢复为 5 个（2 个可删除）、咨询标签恢复为 6 个（2 个可删除）、专用咨询备注记录和删除控件均恢复为 0。

同日完成 `TC-DETAIL-021/022/035` 影像/咨询单正向备注回归：修复后 Run ID `20260827_103720`，`3 passed in 261.33s`，独立只读审计确认事故后影像备注基线保持为 1 条（1 个删除控件）、咨询备注保持为 0 条。首次 Run ID `20260827_102248` 暴露旧版“按纵向最近距离匹配删除控件”缺陷，误删两条运行前影像备注；本轮生成的两条备注已精确清理，但原备注明文未持久化且设备日志无匹配副本，原始 3 条基线尚未恢复。后续实现已改为“精确备注正文 → 直接父行 → 该行唯一删除控件”，并通过真机重跑与摘要审计。

同日继续完成 `TC-DETAIL-012/026/027`：`012` 复用影像备注差集恢复，`026/027` 分别用 `${RUN_TOKEN}` 姓名和运行时生成的 11 位 `${RUN_PHONE}` 验证顾客卡片同步，再恢复八项资料。最终 Run ID `20260827_122044` 为 `3 passed in 308.47s`。首轮 `20260827_111349` 在写入姓名前暴露快照读取个人备注后表单停留底部的问题；增加原生 `scroll_to_top()` 后，Run ID `20260827_112046` 为 `2 passed in 285.41s`。最终 Run 后的独立审计一度被应用明确的“网络异常，请检查网络连接！”阻断；网络恢复后补跑通过，八字段摘要与运行前一致，影像备注保持 1 条/1 个删除控件、咨询备注保持 0 条。

调试时需要保留脚本自动启动的 Appium：

```powershell
.\scripts\run-excel.ps1 -KeepAppium
```

当前平板已有受控顾客和影像数据时：

```powershell
.\scripts\run-excel.ps1 -RunSeeded
```

### Windows 双击运行入口

- 双击根目录的 `run-full-visible.cmd`：使用正常可见的控制台运行，实时显示唯一客户预检、84 条用例的名称、PASS/FAIL 和百分比、失败堆栈、Run ID 与最终状态；结束后窗口会暂停。pytest 输出直接连接控制台并禁用 Python 缓冲，工作簿使用内存模式快速收集；预检和全量阶段都跳过独立 Excel 离线校验。预检通过后显式开启写入与删除门禁，并把全量结果回写 `test_case.xlsx`。首轮完整结束后，只对带 `readonly` 且不含 `mutating`、`destructive`、`persistent`、`no_retry` 的失败用例倒计时 30 秒复跑一次；第二轮使用 `ReadOnlyRetry` 硬门禁且不携带写入、删除或专用顾客预检授权。被安全规则阻止复跑的失败仍保留在最终失败清单，持久化用例 `TC-DETAIL-030` 不会自动再次新增咨询单。第二轮结果覆盖对应只读用例在 Excel 主表中的结果，首轮通过或被阻止复跑的行保持不变；执行记录和 Allure 保留两轮历史。预检、首轮和复跑共用启动器创建的 Appium，结束时只关闭该启动器自己创建的进程。运行前必须关闭 Excel/WPS 中的工作簿。
- 全量状态同时写入 `outputs/launcher-logs/full-run-latest.status.json`；其中 `phase=completed` 表示首轮全部通过，`completed_after_retry` 表示所有可安全复跑的只读失败均已确认通过，`completed_with_failures` 表示仍有失败或存在被安全规则阻止复跑的用例，`preflight_failed` 表示唯一客户预检失败、全量未启动。状态文件同时记录首轮、复跑 Run ID、复跑用例 ID、两轮退出码，以及 `blockedRetryCaseIds`（因写入、删除、持久化或 `no_retry` 标签被阻止复跑）、`remainingFailedCaseIds`（仍未解决）和 `restorationFailed`（恢复失败），`logPath` 指向与控制台内容对应的完整日志。
- Excel 回写时，“实际结果”为 `PASS` 的单元格显示浅绿色，`FAIL` 或 `ERROR` 显示浅红色；下一轮准备阶段写入 `NOT_RUN` 时会清除上一轮遗留的结果颜色。
- 双击根目录的 `run-single-headed.cmd`：保持控制台可见，输入一个精确用例 ID 后运行并回写结果。单条入口跳过独立 Excel 离线校验，并使用内存模式快速收集目标用例，不再因只读随机扫描整张步骤表等待数分钟。所有用例都不再要求输入 `YES`、`DELETE` 或 `PERSIST` 二次确认；写入/删除权限仍根据 Excel 安全标签自动设置，专用客户写入/删除用例仍会自动先执行只读唯一性预检。
- 单条入口也支持命令行参数，例如 `run-single-headed.cmd -CaseId TC-HOME-001 -NoWriteBack`。Android 真机画面始终可见，全量和单条入口现在都会正常显示 Windows 控制台。

只运行某条或某组用例：

```powershell
.\scripts\run-excel.ps1 -CaseId "TC-HOME-001"
.\scripts\run-excel.ps1 -CaseId "TC-HOME-001,TC-HOME-008,TC-HOME-009,TC-HOME-010,TC-HOME-011,TC-HOME-012"
.\scripts\run-excel.ps1 -Tags "smoke,readonly" -RunSeeded
```

以下命令同时适用于当前标准步骤表与旧版单表兼容执行，覆盖已迁移的 Android 登录、首页、顾客列表和顾客详情用例：

```powershell
# 当前测试版已通过的登录回归集（004/005/006 已写入标准步骤表）
.\scripts\run-excel.ps1 -CaseId "TC-LOGIN-001,TC-LOGIN-002,TC-LOGIN-003,TC-LOGIN-004,TC-LOGIN-005,TC-LOGIN-006,TC-LOGIN-007" -NoWriteBack

# 首页 P0 只读导航回归
.\scripts\run-excel.ps1 -CaseId "TC-HOME-001,TC-HOME-008,TC-HOME-009,TC-HOME-010,TC-HOME-011,TC-HOME-012,TC-CASE-001" -NoWriteBack

# 病例库 P2 只读负向搜索：不存在标签由运行变量生成，不依赖病例种子
.\scripts\run-excel.ps1 -CaseId "TC-CASE-002" -NoWriteBack

# 病例库 P1 只读回归：依赖存在“火”标签及对应病例
.\scripts\run-excel.ps1 -CaseId "TC-CASE-003,TC-CASE-004" -RunSeeded -NoWriteBack

# 影像阅览 P1 只读回归：唯一匹配“咨询单特定”；030 只校验五项评分非空，不采集或记录具体评分值
.\scripts\run-excel.ps1 -CaseId "TC-IMAGE-030,TC-IMAGE-032,TC-IMAGE-033,TC-IMAGE-034" -RunSeeded -NoWriteBack

# P1 新增咨询单：动态选择首张未关联影像，保存后永久保留新增咨询单，不删除或回滚
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-030" -RunSeeded -AllowMutation -NoWriteBack

# 影像/咨询单备注正向回归：创建本轮唯一备注，断言后由恢复会话只删除该行
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-012,TC-DETAIL-021,TC-DETAIL-022,TC-DETAIL-035" -RunSeeded -AllowMutation -CustomerPreflightVerified -NoWriteBack

# 自包含删除回归：只创建并删除本轮唯一标签/备注，仍需显式开启全部门禁
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-020,TC-DETAIL-037" -RunSeeded -AllowMutation -AllowDestructive -CustomerPreflightVerified -NoWriteBack

# 顾客列表 P0：需要受控顾客种子数据
.\scripts\run-excel.ps1 -CaseId "TC-CUSTOMER-020,TC-CUSTOMER-021" -RunSeeded -NoWriteBack

# 顾客详情资料校验：仅在本地专用顾客完成唯一匹配只读预检后执行
# 003 校验 33 字符姓名长度；004/005 校验男/女性别；007 校验生日；008~010 校验地址；013/014 校验个人备注；015 校验婚姻状态保密
# 每条资料写入用例均在内存中保存八项原值，退出时恢复并重新打开编辑页精确比对
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-001,TC-DETAIL-003,TC-DETAIL-004,TC-DETAIL-005,TC-DETAIL-007,TC-DETAIL-008,TC-DETAIL-009,TC-DETAIL-010,TC-DETAIL-011,TC-DETAIL-013,TC-DETAIL-014,TC-DETAIL-015" -RunSeeded -AllowMutation -CustomerPreflightVerified -NoWriteBack

# P1 顾客卡片同步：保存本轮唯一姓名/手机号，返回列表精确断言后恢复八项资料
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-026,TC-DETAIL-027" -RunSeeded -AllowMutation -CustomerPreflightVerified -NoWriteBack

# P1 影像/咨询单标签空值与空格校验：点击保存后校验提示，再取消弹窗；不创建或删除标签
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-017,TC-DETAIL-019,TC-DETAIL-032,TC-DETAIL-034" -RunSeeded -NoWriteBack

# P1 影像/咨询单标签创建与特殊字符：每条用例创建本轮唯一标签，断言后仅删除差集新增项并验证原集合恢复
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-016,TC-DETAIL-018,TC-DETAIL-031,TC-DETAIL-033" -RunSeeded -AllowMutation -CustomerPreflightVerified -NoWriteBack

# P1 顾客详情影像管理态：只切换管理/完成，不勾选或删除影像
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-024" -RunSeeded -NoWriteBack

# P1 顾客详情备注空格校验：输入空格后真实点击提交，并验证备注记录数量不增加
# 按钮视觉/原生 enabled 状态不作为成败依据，以实际未新增记录为准
.\scripts\run-excel.ps1 -CaseId "TC-DETAIL-023,TC-DETAIL-036" -RunSeeded -NoWriteBack

# P1/P2/P3 首页搜索（已写入标准步骤表）：002/003 为空结果；004/005/006 需要受控顾客种子；007 需要本地 SEEDED_CUSTOMER_PHONE
.\scripts\run-excel.ps1 -CaseId "TC-HOME-002,TC-HOME-003,TC-HOME-004,TC-HOME-005,TC-HOME-006,TC-HOME-007" -RunSeeded -NoWriteBack

# P1 顾客列表搜索：001~004 已写入标准步骤表；空结果、姓名和手机号搜索均为只读
.\scripts\run-excel.ps1 -CaseId "TC-CUSTOMER-001,TC-CUSTOMER-002,TC-CUSTOMER-003,TC-CUSTOMER-004,TC-CUSTOMER-005,TC-CUSTOMER-006" -RunSeeded -NoWriteBack

# P1 顾客性别/年龄筛选（已写入标准步骤表）：使用真实 Android resource-id 选项；011 暂不执行
.\scripts\run-excel.ps1 -CaseId "TC-CUSTOMER-009,TC-CUSTOMER-010,TC-CUSTOMER-012,TC-CUSTOMER-013,TC-CUSTOMER-014,TC-CUSTOMER-015,TC-CUSTOMER-016" -RunSeeded -NoWriteBack

# P1 顾客筛选重置：只设置临时条件并点击重置，不点击“确定”或提交筛选
.\scripts\run-excel.ps1 -CaseId "TC-CUSTOMER-022" -RunSeeded -NoWriteBack

# P1/P2 顾客日期与联合筛选（007/008/017/018 均已写入标准步骤表）：使用真实 Android 双滚轮
.\scripts\run-excel.ps1 -CaseId "TC-CUSTOMER-007,TC-CUSTOMER-008,TC-CUSTOMER-017,TC-CUSTOMER-018" -RunSeeded -NoWriteBack
```

单表中的 CSS 定位器和 URL 断言不会直接用于原生 Android；执行器会将已支持的用例映射为真实 resource-id、Activity 和原生页面源断言。`TC-LOGIN-006` 的 `login_protocol_tv` 是包含局部 `ClickableSpan` 的不可点击 `TextView`，普通元素中心点击不会命中蓝色《用户协议》链接；步骤通过 `click` 的元素内相对坐标 `0.84,0.5` 点击链接，并校验 `.WebViewActivity` 与标题“用户协议”。`TC-HOME-006` 使用手机号片段 `186` 做首页只读搜索，进入顾客列表后校验卡片存在，并对首个可见 `a_records_unique_tv` 执行 `text_contains`；不会在日志或报告中采集完整手机号。`TC-HOME-007` 使用 `${SEEDED_CUSTOMER_PHONE}` 做首页只读精确搜索，要求 `a_records_cl` 数量严格等于 1 且 `a_records_unique_tv` 与运行时变量完全相等；该变量按敏感信息脱敏；2026-08-31 已用唯一受控顾客完成真机通过，新的环境仍必须在本地配置该变量。`TC-CUSTOMER-020` 将原生卡片作为可点击的详情入口，并用图像控件验证性别/头像类字段；搜索切片使用 `empty_tv` 或顾客卡片数量断言结果，性别/年龄筛选使用真实页面的 `customer_records_other_sex_*_tv` 和 `customer_records_other_age_*_tv` 控件，并以顾客卡片数量断言结果。真实 Android 页面没有“未知”性别选项，因此 `TC-CUSTOMER-011` 明确保留在待补清单，不会误报自动化通过；真机已确认 `TC-CUSTOMER-015` 的“45岁以上”筛选显示 `empty_tv=暂无内容`。日期范围筛选使用真实页面的 `customer_records_date_*` 控件和双组 `date_picker_*_wheel`，由 `select_date_range` 根据当前日期文本计算滚轮步数并在确认后校验 `customer_records_date_tv`；Android 当前可选日期起点为 `2015-02-04`，`TC-CUSTOMER-007` 使用 `2015-02-04~2015-02-13` 校验纯日期空结果，`TC-CUSTOMER-017` 使用 `2015-02-04~2015-02-28` 校验日期、性别和年龄联合空结果。`TC-CUSTOMER-022` 已按真机备注占位文本“搜索备注记录”和 `selected` 属性接入：设置性别、年龄与临时备注后点击重置，逐项校验恢复且保持筛选面板打开，全程不点击“确定”。`TC-DETAIL-024` 使用真机确认的 `customer_detail_manager_tv`、`a_records_detail_all_select_ifv` 和 `customer_detail_delete_tv` 验证管理态切换，全程不勾选或删除影像；`TC-DETAIL-023` 和 `TC-DETAIL-036` 使用 `a_records_detail_all_remark_ifv`、`a_consultation_result_remark_ifv`、`records_remark_remark_et` 与 `records_remark_remark_commit_tv`，通过 `${SPACE}` 输入单个空格后真实点击提交，并持续校验备注记录数量及当前 Activity 不变；按钮亮度或原生 `enabled` 属性不作为功能结果。`TC-IMAGE-030` 在 `.SkinResultActivity` 内使用真机唯一匹配的名称/评分 XPath，分别校验炎敏、肤色、色素、亮度和红区五项评分控件可见且非空；运行日志、工作簿和报告均不采集或记录具体评分值。`TC-IMAGE-032/033/034` 先以“咨询单特定”完成唯一顾客匹配，再分别校验 `f_skin_result_info_view_case_tv` 跳转 `.CaseActivity`、报告菜单跳转 `.SkinReportActivity`，以及 `cover_prompt_v1_tv=不保存直接退出` 后返回 `.CustomerDetailActivity`；当前报告页没有旧用例描述中的“完成”按钮，已按真机存在的 `skin_report_pdfv` 和分享入口修正断言。三条路径均不保存或修改顾客、影像、报告数据。`TC-DETAIL-029` 已通过瞬时页面文本断言“请选择需要删除的影像”，全程不勾选或删除影像。`TC-DETAIL-001/003/004/005/007/008/009/010/011/012/013/014/015/026/027/030` 会触发保存动作，因此带有 `mutating` 安全标签；`TC-DETAIL-006` 的真机性别选择器只有男、女，没有旧用例要求的“保密”选项，已标记为 `BLOCKED_PRODUCT` 且不生成执行步骤。旧版单表工作簿未选择用例时仍会拒绝执行；当前含步骤表的工作簿只装载已接入 Android 步骤的用例。

2026-09-02 更新：上段关于 `TC-DETAIL-023/036` 的按钮状态判定已被替换。两条用例现在对提交控件执行真实中心点击，并在观察窗口内持续校验 `a_records_remark_list_content_tv` 数量及当前 Activity 均不变化；按钮亮度或原生 `enabled` 属性不再决定结果。`TC-DETAIL-033` 使用本地 `YANJIA_MUTATION_CUSTOMER_QUERY` 在顾客搜索栏精确匹配卡片手机号，并等待咨询单备注入口加载后再进入写入恢复会话。

2026-08-27 更新：`TC-DETAIL-029` 已按当前 Android 页面改为直接文本断言。用例唯一匹配“咨询单特定”种子顾客，进入管理态但不勾选影像，点击删除后对页面中心实际文本“请选择需要删除的影像”执行 `page_source_contains`；该提示是瞬时节点，避免对节点继续查询可见性导致 stale element。随后退出管理态，全程不修改业务数据。`TC-DETAIL-030` 已按真实完成路径接入：先比较历史影像检测时间与现有咨询单检测时间，动态打开首张未关联影像；返回时断言 `cover_prompt_v2_tv=保存并退出`，保存后校验所选检测时间出现在咨询单卡片中。该用例是持久化写入，按用户要求保留新增咨询单，不执行删除或回滚；若当前可见影像均已关联，则在写入前安全失败。同日 Android 16/API 36 平板正式回归通过：`1 passed in 251.00s`。

`TC-DETAIL-003` 在专用顾客编辑页输入 33 字符姓名并校验“客户名称长度不合法”，随后在恢复上下文中强制写回运行前姓名、手机号、性别、生日和地址，再重新打开编辑页逐字段比对；专用查询和原资料均不会写入日志或报告。

`TC-DETAIL-004/005` 使用真机确认的 `customer_edit_sex_tv`、`customer_edit_sex_man_tv` 和 `customer_edit_sex_female_tv` 保存目标性别，随后重新打开编辑页执行 `text_equals`。`TC-DETAIL-007` 使用 `customer_edit_birthday_tv` 和单组 `date_picker_*_wheel` 选择生日；原步骤写“2022-08-03”，但输入数据和验证点均为“2022-08-04”，已统一为后者。`TC-DETAIL-008/009/010` 使用 `customer_edit_address_et`，真机分别确认空地址可保存、特殊字符原样保存、单空格保存后归一为空。`TC-DETAIL-013/014` 先对顾客资料 `ScrollView` 执行原生滚动，再通过 `customer_edit_remark_et` 校验空备注和特殊字符备注。`TC-DETAIL-026/027` 先保存本轮唯一姓名或手机号，返回 `.CustomerRecordsActivity` 后分别对 `a_records_name_tv`、`a_records_unique_tv` 执行 `text_equals`，恢复时可从当前唯一筛选卡片或未变的资料键重新进入编辑页。以上资料写入用例的恢复上下文都会强制写回运行前姓名、手机号、性别、生日、邮箱、婚姻状态、地址和个人备注，并再次重开逐字段比对。`TC-DETAIL-016/018/031/033` 使用 `${RUN_TOKEN}` 创建本轮唯一影像或咨询单标签；`TagMutationSession` 不假设应用保存后的显示文本等于原输入，而是比较保存前后的标签多重集合，只允许恰好一个新增项，再按新增显示文本和最近删除控件精确清理，最后验证原标签文本、数量和可删除控件数量全部恢复。`TC-DETAIL-006` 因产品未提供“保密”性别选项保持阻塞。

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
- 专用顾客用例（顾客资料写入与卡片同步、`TC-DETAIL-016/018/020/031/033` 标签创建恢复、`TC-DETAIL-012/021/022/035/037` 备注创建恢复或删除，以及 `TC-DETAIL-017/019` 影像标签校验）必须配置本地 `YANJIA_MUTATION_CUSTOMER_QUERY`，该值按敏感信息处理，不会出现在配置或目标对象的 `repr` 中。
- 每次顾客写入运行还必须在刚完成只读唯一匹配预检后显式提供 `--customer-preflight-verified`（PowerShell 为 `-CustomerPreflightVerified`）；该授权不应长期保存在 `.env`。
- `TC-DETAIL-001/003/004/005/007/008/009/010/011/013/014/015/026/027` 已接入“专用查询—唯一匹配—内存快照—强制恢复—重新打开验证”安全执行流；快照和恢复覆盖姓名、手机号、性别、生日、邮箱、婚姻状态、地址和个人备注，缺少专用查询或本次只读预检标记时，会在创建 Appium driver 前记录 `SKIP`。
- `CustomerMutationSession` 只允许一个 mutation 回调；回调成功或异常都会恢复快照。恢复抛错或验证失败会升级为 `CustomerRestoreError`，提示停止后续写入用例。
- `TagMutationSession` 同样只允许一个写入回调；回调成功或异常都会执行标签集合差集恢复。若原标签缺失、差集不是恰好一个新增项或恢复后集合不一致，会升级为 `TagRestoreError` 并停止后续写入。
- `TC-DETAIL-020` 在 `TagMutationSession` 内创建本轮唯一标签，再通过 `delete_current_run_tag` 只删除集合差集中的唯一新增项；动作不接受 Excel 定位器，防止用例绕过运行器的配对和恢复校验。
- `RemarkMutationSession` 为 `TC-DETAIL-012/021/022/035/037` 按影像或咨询单范围记录原备注文本、记录数和删除控件数；恢复时通过精确正文的直接父行定位唯一删除控件，禁止按屏幕距离推断。`TC-DETAIL-037` 的 `delete_current_run_remark` 同样只允许删除本轮唯一新增备注。异常退出会强制恢复，失败会升级为 `RemarkRestoreError`。
- `TC-DETAIL-023`、`TC-DETAIL-036` 输入单个空格后真实点击提交，并持续校验备注记录数量不增加且仍停留在备注编辑页；应用拒绝空格备注，因此不需要开启写入门禁。
- `TC-DETAIL-017/019/032/034` 只验证空值或单空格标签提示并取消弹窗；真机探针已确认标签集合保持不变，不创建或删除标签，也不需要开启写入门禁。
- `TC-DETAIL-030` 是显式授权的持久化写入用例：需要 `--run-seeded --allow-mutation --no-excel-writeback`，动态选择未关联影像，保存后保留新增咨询单，不执行删除或回滚；不需要专用顾客资料预检标记。
- 当前启用 `TC-DETAIL-012/021/022/035` 四条可恢复备注写入用例和 `TC-DETAIL-020/037` 两条自包含删除用例；它们只删除本轮差集新增项，原有标签、备注和影像均不属于允许删除目标，其他破坏性用例仍保持 `是否执行=否`。
- 任意 `CustomerRestoreError`、`TagRestoreError` 或 `RemarkRestoreError` 会锁住当前 pytest 会话的后续非只读用例：在创建 driver 前记为 `SKIP`，原失败记为 `ERROR`，并通过 `restoration_failed=true` 记录该问题。同轮熔断仍保留；这些写入相关失败和受影响用例不会自动复跑，必须先人工审计并恢复数据。
- 首次 Appium 会话创建和重建均纳入错误记录；可恢复的创建错误只允许只读用例重试一次，最终失败写回 `ERROR`。
- 默认捕获模式下，pytest 终端堆栈、捕获输出和 Allure 用例/步骤/fixture 错误在报告输出前统一脱敏，覆盖已登记敏感值及其 repr/JSON 转义形式，保留异常类型和定位信息。
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

第二条命令会替换当前 85 个内置目录用例的步骤；正常日常执行不需要运行它。

## 本地质量检查

```powershell
.\.venv\Scripts\python.exe -m pyright --pythonpath .venv/Scripts/python.exe
.\.venv\Scripts\python.exe -m ruff check src tests scripts
.\.venv\Scripts\python.exe -m pytest tests\unit -q
```

## 已知边界

- 影像阅览首次进入可能需要下载较大资源，应放在扩展套件，不阻塞最短 Smoke。
- 影像备注提交按钮可能以较低亮度显示，咨询单备注按钮在输入空格后可能显示为启用，但两者实际都不会新增空格备注；`TC-DETAIL-023`、`TC-DETAIL-036` 以提交后的记录数量不增加为功能判定。
- Run ID `20260827_102248` 曾因旧的备注行删除配对算法误删专用顾客两条运行前影像备注；生成数据已清理，修复后的 Run ID `20260827_103720` 与事故后摘要审计通过，但原始 3 条备注基线仍缺 2 条。补齐专用测试数据前，不应把事故后的 1 条基线描述为原始数据已恢复。
- Run ID `20260827_122044` 完成后，顾客列表曾短暂显示“网络异常，请检查网络连接！”，独立审计在目标查询前安全停止；网络恢复后八字段与备注双摘要审计均已补跑通过。
- 搜索、筛选和编辑依赖受控测试数据；顾客必填校验、标签创建和两条自包含删除用例已接入专用查询、唯一匹配、写后恢复和真机回归，影像删除等非自包含破坏性流程仍未启用。
- 颜佳AI当前页面均能从原生 accessibility 树取得稳定 `resource-id`，不使用 CSS 或 URL 断言。
- Appium 仅监听 `127.0.0.1`，不要启用全局 `--relaxed-security`。
- 带 `requires_auth` 且不是登录模块的用例如果失败现场为登录页，会先保留原始失败证据，再自动登录并恢复首页，防止后续用例级联失败；可见全量启动器在首轮结束 30 秒后只复跑可重复执行的只读失败，不继承写入或删除授权。底层 Appium 会话/连接故障同样只对安全的 `readonly` 用例执行一次即时基础设施重试。
- Allure 报告保留历史趋势，并记录平板型号、Android/SDK 和颜佳AI版本；失败会按业务断言、定位等待、基础设施和配置分类。
