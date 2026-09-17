# Windows Python 运行依赖与正式发布

`config/runtime-requirements.lock` 是需要评审、提交且进入签名运行包的依赖清单。每一项都有精确版本和一个 Windows x64 / CPython 3.12 兼容 wheel 的 SHA-256；包括直接依赖、间接依赖和 pip。正式流水线只安装这份清单，不在发布时重新解析浮动版本。

锁文件放在 `config/`，因此现有 TOP 白名单导出器会将它纳入 package digest，无需修改或重新发布 TOP 适配器。

## 新版本安装

核验正式 Release 的 tar.gz、签名、provenance 和证据摘要后，解压到新的版本目录。从其中的 `android/` 执行：

```powershell
.\scripts\install-runtime.ps1 -Python C:\Path\To\Python312\python.exe
.\.venv\Scripts\python.exe -I .\scripts\runtime_dependencies.py
```

安装器使用 `--require-hashes --only-binary=:all:`，只接受锁定 wheel。默认创建本版本的 `.venv`，若目标已经存在则拒绝操作，避免改动旧 Worker 的依赖。失败目录保留用于诊断；另选新的版本/尝试目录重试。

安装器只将当前解压版本的 `src/` 加入新环境，不链接开发仓库的旧环境，也不需要未锁定的 editable/build 依赖。锁文件通过 `.gitattributes` 固定 LF，保证 Windows 证据 job 与 Linux 签名 job 计算同一份字节摘要。

TOP 通过 `run-excel.ps1 -IgnoreDotEnv` 调用时，会先验证 Windows x64 CPython 3.12、锁文件摘要、全部已安装 distribution 的集合与版本、项目直接依赖约束和 `pip check`。依赖缺失、版本改变或安装了额外插件均失败关闭。`-SkipPreflight` 和 `-SkipExcelValidation` 不会跳过这个检查。

校验针对安装包集合和版本，安装过程负责 wheel 哈希校验；它不是逐文件防篡改检查。旧版本安装目录、旧 `.venv`、旧 Release 资产和包记录应全部保留。已有设备数据副作用仍按测试授权和唯一顾客预检规则处理。

## CI 发布链路

1. Windows job 用干净虚拟环境安装哈希锁定的 Python 依赖，产生 `runtime-inventory.json`，并执行无设备单元测试。
2. Syft 扫描该环境的 `Lib/site-packages`，生成含 Python PURL 与版本的 `sbom.spdx.json`。扫描目标不是只有源码的打包目录。
3. Ubuntu 签名 job 下载同一 workflow / commit 的 Windows 证据，使用 Trivy `--list-all-pkgs` 扫描 SBOM，保留所有严重等级的漏洞。
4. `verify_release_evidence.py` 核对 commit、锁文件摘要、Windows 安装清单、每个锁定依赖的 SBOM 版本及每个 SBOM Python 组件的扫描版本。空 SBOM、空扫描结果、缺包和版本不一致均阻断发布。
5. 保持既有 CRITICAL 漏洞阻断策略；HIGH/MEDIUM/LOW/UNKNOWN 会保留在报告中，不能将 CRITICAL 门禁通过表述为“没有漏洞”。
6. 四份既有 Sigstore 签名继续覆盖包、provenance、SBOM 和漏洞报告。新增锁文件、安装清单和覆盖报告的摘要进入签名 provenance 的 `runtimeDependencies`；Release manifest 同时列出这些摘要。
7. Release 包内锁文件与单独发布的锁文件必须逐字节相同。下载端还须核对签名 provenance 的三个依赖证据摘要，不能只信未单独签名的 manifest。

Release 在原有 11 个资产基础上新增：

- `runtime-requirements.lock`
- `runtime-inventory.json`
- `dependency-coverage.json`

该 SBOM 覆盖 Python 自动化运行依赖。Windows 系统、TOP 控制器环境、Node/Appium 服务、ADB 和设备应用属于独立运行设施，本清单不宣称覆盖它们。

## 更新锁文件

在 Windows x64 CPython 3.12 下，为一次更新使用全新的空 wheelhouse，显式传入经过评审的直接依赖版本和 pip。示例中每个占位版本必须先替换为实际选定版本：

```powershell
python -m pip download --only-binary=:all: --dest C:\Temp\new-runtime-wheels `
  'Appium-Python-Client==<version>' 'allure-pytest==<version>' `
  'openpyxl==<version>' 'python-dotenv==<version>' 'pytest==<version>' 'pip==<version>'
python scripts/build_runtime_lock.py --wheelhouse C:\Temp\new-runtime-wheels `
  --output config/runtime-requirements.lock
```

审查锁文件变化后，用安装器创建新环境，再执行单元测试、SBOM 和扫描覆盖检查。不要修改已签名 Release，也不要让运行中的旧环境原地升级。

`pyproject.toml`、`testops-package.json` 与新 Git 标签必须版本一致。只有新正式 Release 供应链验证通过、TOP 完整 84 条验证 Run 最终 PASSED 后才激活；单元测试或 CI 成功不能代替真机验收。
