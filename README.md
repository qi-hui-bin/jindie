# 银行流水转精斗云凭证

银行覆盖、使用边界及验证结果见 [CURRENT_SCOPE.md](./CURRENT_SCOPE.md)；开发规格见 [PLAN.md](./PLAN.md)。当前源码支持从本地页面通过 Excel 新增账套，也兼容既有 JSON 账套。流水可逐张复核、手工制证、跳过及撤回、同日凭证拆并、导出留存和单张修订重导。

## Windows 双击启动

完整解压 `Jindie-Windows-x64-multibook.zip`，不要在压缩包里直接运行。Windows x64 上不需要安装 Python 或 Excel。

- **先试用**：双击 `Start-Demo.cmd`，上传同目录 `demo` 下的 `boc-voucher-synthetic.xls`，期间选 `2026-11`。演示不会写入正式数据库，也不能导出实际导入文件。
- **正式使用**：用当前源码构建的启动包，双击 `Start-Jindie.cmd` 后打开「配置 → 新增账套」。已有账套 JSON 仍从 `%LOCALAPPDATA%\Jindie\books` 加载，无需重新录入。
- 启动器默认打开系统浏览器，仅在本机监听。没自动打开时，把窗口里显示的地址复制到 Chrome 或 Edge。关闭启动窗口会停止服务，已保存的批次不会丢失。
- 升级时解压到新目录即可，配置和数据库仍留在 `%LOCALAPPDATA%\Jindie`。迁移前会备份数据库；若启动报告迁移失败，请保留原库和 `.bak`，不要覆盖它们。

**重新构建启动包（维护人员）**：在 Windows 上准备 [Python.org 发布的 3.12 x64 嵌入式 ZIP](https://www.python.org/downloads/release/python-31210/)，安装下节的构建依赖后，在项目目录运行：

```powershell
.\.venv\Scripts\python.exe -m tools.build_windows --runtime-zip .\build\python-3.12.10-embed-amd64.zip
```

默认生成 `dist\Jindie-Windows-x64.zip`；脚本不会覆盖已有的 ZIP。也可用 `--runtime-sha256` 核对运行时归档的预期哈希。

### 从页面新增账套

1. 在「配置 → 新增账套」下载主数据填写表，填写「科目」「辅助核算」「银行账户」。按文本填写账号及编码，保留前导零。可一次绑定多个账户，各账户可选择内置中行、建行、兴业格式或自动识别。
2. 上传填写表和同一目标精斗云账套的空白凭证导入模板，填写凭证字、人民币币别代码和编号类型，核对无误后勾选财务核验。程序会检查表头、编码引用和模板是否空白，但无法自动确认两个文件确实来自同一个精斗云账套。
3. 创建后不用重启即可上传样本。页面新建的账套默认没有自动匹配规则，可以先手工制证，也可以在核对样本后配置自动制证规则。

更新页面账套时，打开「配置 → 更新账套配置」，下载当前主数据，在 Excel 中修改后上传**完整工作簿**。可替换空白模板，也可留空沿用。已有账户标识不能改绑另一账号；停用账户时可从新工作簿移除该行，不会删除历史。新版只用于新批次，相同流水文件在解析条件未变时仍打开原批次。旧批次保留原模板和编码，不会因发布新版自动重算。

页面建立的账套和模板保存在本机 SQLite 数据库中，备份数据库时也应按财务资料保管。已有 JSON 账套继续按原配置读取。

### 接入新的银行流水格式

在页面账套的「配置」中，打开对应银行账户旁的「配置流水格式」。按样本原表头填写映射，支持单列正收负支或收支分列；要求本方完整账号、流水号、日期及金额。没有币种列时，只有确认该格式仅包含人民币才可继续。先上传 `.xls` 或 `.xlsx` 样本核对行号、日期、收支笔数和金额。样本无异常后，还须勾选确认才能发布；预览不会创建流水批次。

发布只绑定当前账户，不会改动其他账户或历史批次。已发布映射可在同一账户的后续上传中复用；更新账套主数据时保留「自定义」即可沿用，改为「自动」可恢复内置格式识别。当前不支持用页面配置仅有表头账号、只有正数且另列标记收支、自定义币种或需要过滤非交易行的格式；此类格式仍由维护人员适配。

### 配置自动制证规则

打开页面账套的「配置 → 自动制证规则」，或从复核批次打开「用本批流水试跑规则」。限定一个银行账户及收支方向，填写精确对方户名、银行字段条件，或同时填写两项；指定对方科目、辅助编码和摘要。选同一账户的已保存批次试跑。系统使用批次中能解析且不重复的源流水，不应用该批次后续的手工修订或跳过，也不写回凭证。

没有命中、命中行与现有规则冲突、或不能生成有效草稿时不能发布。核对命中范围和分录后确认发布。发布、修改或停用规则都只影响新批次；同一文件重传不会自动用新规则重算。页面规则当前只支持普通一借一贷，生成的草稿必须人工复核；固定多分录、合并和高级条件仍由维护人员通过原有 JSON 配置处理。停用规则后才可从后续主数据版本移除它引用的账户或编码。

### 查看待确认清单

在左侧「待确认清单」按账套、期间和问题类型筛选。清单只收录**已经完成解析和查重、但遇到制证问题**的流水，可以定位到原批次、原行，或按原配置手工制证。生成手工草稿或明确跳过后，该行从清单移除；撤销跳过会恢复未处理项。手工草稿仍须复核，新增编码或规则不会自动重算旧批次。解析失败、重复和已生成草稿的流水继续在各批次查看。

页面账套发布新配置后，可对清单中的合格流水逐条点击「用新版配置试算」。试算只重新匹配制证规则，不重新解析流水或修改原凭证日期与金额。核对分录后，单条转入独立复核批次，使用新版配置和模板，原批次保留来源及跳转链接，其他草稿和既有导出不变。未成功试算的流水继续留在清单。纯 JSON 账套不支持此操作。

## Windows 开发验证

使用 Python 3.12，在项目目录运行：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-test.lock
.\.venv\Scripts\python.exe -m unittest discover -v
```

以下示例生成并读取合成流水，可分别检查导入汇总和已保存批次：

```powershell
.\.venv\Scripts\python.exe -m tests.fixtures .\build\sample-data
.\.venv\Scripts\python.exe -m jindie --db .\build\demo.sqlite3 import .\build\sample-data\boc-synthetic.xls --book demo --account-id boc --account-number 0000000000009130 --period 2026-11
.\.venv\Scripts\python.exe -m jindie --db .\build\demo.sqlite3 inspect 1
```

样本生成命令不会覆盖已有文件。合成中行样本含 49 笔（37 收、12 支），建行样本含 91 笔（32 收、59 支）；这些数据不代替财务提供的脱敏底稿。

- 自动识别内置银行格式；多张表匹配时用 `--sheet` 选择。可用 `--profile` 指定内置格式，或用 `--profile-json` 加载自定义字段映射；配置结构参见 [bank_profiles.py](./jindie/bank_profiles.py)。
- 建行 0306 支持 `YYYYMMDD HH:MM:SS` 交易时间。若之前用旧版导入过并收到日期异常，请用新版重新上传原文件。交易日和记账日不同时仍会保留异常；核实口径后可使用 `--posting-preference transaction` 或 `--posting-preference posting`。账号被遮蔽或缺失时，仅在核实后使用 `--confirm-binding`。
- 不提供 `--db` 时，数据库位于 `%LOCALAPPDATA%\Jindie\statements.sqlite3`。数据库保存必要的原始流水字段。实际账户不要通过命令行参数传入；省略 `--account-number` 后由终端提示输入。请勿将真实工作簿或数据库加入仓库。

## 贯通试算：银行流水 → 25 列工作簿

以下命令使用独立的合成数据目录。在新数据库中，这个导入批次的 ID 为 `1`；使用已有数据库时以 `import` 返回的 `batch_id` 为准。

```powershell
.\.venv\Scripts\python.exe -m tests.voucher_fixtures .\build\voucher-demo
.\.venv\Scripts\python.exe -m jindie --db .\build\voucher-demo.sqlite3 import .\build\voucher-demo\boc-voucher-synthetic.xls --book demo --account-id boc --account-number 0000000000009130 --period 2026-11
.\.venv\Scripts\python.exe -m jindie --db .\build\voucher-demo.sqlite3 preview 1 --config .\build\voucher-demo\voucher-config.json --template .\build\voucher-demo\synthetic-template.xlsx --output .\build\voucher-demo\preview.xlsx --first-number 101 --confirmed-skips .\build\voucher-demo\confirmed-skips.json --include-partial
```

合成回归应输出 44 张凭证、91 条分录、2 笔已核实跳过。11 月 12 日的两笔手续费合并为借方 9.90 元、银行贷方 9.00 元和 0.90 元。试算表页名为「试算预览（不可导入）」，仅供查看，请勿上传精斗云。

`voucher-config.json` 同时提供版本化的账套主数据、银行科目和匹配规则。规则显式设为 `enabled: true` 才参与匹配；不匹配、规则冲突、缺编码或原始行异常会在报告中列出。存在排除项时，只有指定 `--include-partial` 才写工作簿。`--confirmed-skips` 接收同批已核实的流水号及原因，不根据「划转」字样自动跳过。`--used-number` 可以重复传入，供试算检查既用号；试算不占用正式凭证号。

## 启动本地浏览器界面

安装上述依赖后，生成合成数据（已有数据目录可跳过第一条）：

```powershell
.\.venv\Scripts\python.exe -m tests.voucher_fixtures .\build\voucher-demo
.\.venv\Scripts\python.exe -m jindie --db .\build\browser-demo.sqlite3 serve --config .\build\voucher-demo\voucher-config.json
```

在 Chrome 或 Edge 打开 [http://127.0.0.1:8765](http://127.0.0.1:8765)。上传合成的 `boc-voucher-synthetic.xls`，期间选 `2026-11`。可在历史批次恢复复核。合成账套默认禁用正式导出；需要查看合成 Excel 时，使用上面的 `preview` 命令。

复核页顶部汇总收支，下方按凭证显示状态、来源和分录；异常与跳过记录可展开查看。宽屏下导出在右侧，窄屏下可用顶部「导出 Excel」跳到导出面板。

本地配置在原有 `ledger` 和 `rules` 之外添加 `web` 对象：`label` 为显示名称，`accounts` 为账户列表（每项包含 `id`、`number`、可选的 `label` 和 `posting_preference`），`template` 为模板路径（可相对于配置文件）。可重复指定 `--config` 登记多个账套。真实配置文件、模板和数据库请保存在仓库之外。

正式导出前，财务须核验目标账套的科目及辅助核算，并在 `ledger` 中设置 `finance_verified: true`；在 `web` 中登记核验后的空白模板 `verified_template_hash`（`Get-FileHash -Algorithm SHA256` 的小写哈希）、凭证字 `voucher_word`、人民币币别 `currency_code` 和编号是否按文本输出 `text_number`。旧批次保留建立复核时的科目和规则，不会因启动配置改动自动更新。导出时填写起始号及精斗云既用号；本地已分配的编号也会参与检查。导出成功后可重复下载原文件，确认精斗云导入结果须由操作员手动登记。

跳过流水时，填写原因及核对依据（例如既有凭证号）；未导出的跳过可撤回。合并凭证只接受同批、同日、同账户、同币种的记录；如果只跳过合并凭证中的一笔，先按来源流水拆分。拆并后的凭证须重新复核。

对于已成功解析、无查重问题的人民币流水，如果未匹配规则、规则冲突或缺少科目或辅助核算，可在复核页的异常列表点击「手工制证」。核对源行后选择对方科目、填写必要的辅助编码和制证依据。程序按源流水生成银行分录并检查借贷，草稿须另行复核；手工处理不会修改长期规则。

精斗云报告导入失败时，先逐张登记结果并在精斗云核实该张凭证确实未入账。登记为「已核实未导入」后，可从原导出记录发起单张修订。修改、重新复核并使用新凭证号单独导出；原文件、旧号码及同批其他凭证不变。不要重新导入原来整批的 Excel。

## 坤昇 SOP 使用说明

- 本版可以直接识别中行带方括号英文说明的表头，以及兴业人民广场文件的 19 列格式。兴业用「记账日期」、以「唯一流水编号」查重。若另一个兴业账户的格式有差异，可在该本地账户的 `web.accounts[]` 中提供 `bank_profile` 对象，字段与命令行 `--profile-json` 一致。
- 每次正式导出同时保存两份文件。「精斗云导入」只有一个「凭证模版」工作表，上传给精斗云。「核对底稿」有七页，在同一个导出详情页下载，用于复核原始来源、分录、科目汇总和存疑项。旧版已生成的下载记录保留原文件。
- 常规规则可用 `counterparty_account` 条件精确匹配对方账号。例如按已核实的内部转账账户配置对方银行科目。已经入账的对应流水仍由操作员填写原因及凭证依据后跳过。
- 工资及税费的具体金额由财务核实后按批次录入规则。固定多分录必须配置明确的账户和同日流水号，系统按流水金额生成银行分录，按 `fixed_entries` 中明确的金额生成其余分录。只有所有指定流水匹配成功且借贷平衡，才产生待复核凭证。每笔实发银行流水保留一行。

下面是合成工资规则示例。实际使用时替换账套、账户、来源流水号、科目、辅助编码和金额，金额以分为单位。例中实发 100.00 元，对应应发 120.00 元及个税 20.00 元。社保、公积金或其他已确认款项可追加分录，并在对应项中添加 `auxiliary` 数组，例如 `[{"category":"其它","code":"010"}]`。

```json
{
  "id": "payroll-example",
  "version": 1,
  "book_id": "demo",
  "account_id": "cib",
  "direction": "expense",
  "summary": "发放工资",
  "merge": "bank_detail",
  "enabled": true,
  "bank_references": ["000001", "000002"],
  "fixed_entries": [
    {"subject": "221101", "side": "debit", "amount_minor": 12000},
    {"subject": "222112", "side": "credit", "amount_minor": 2000}
  ]
}
```

税费拆分使用相同结构，`bank_references` 填一笔已核实的缴税流水，各税费科目按确认金额分别列为借方。配置更新后新批次使用新规则，已创建的复核批次沿用原配置快照。
