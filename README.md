# 区域车辆录入工具 vehicle-flow

面向**汽车经销商 / 进口商**的「入库 / 末销 / 看板」三件套。一个 Python 脚本 + 一个本地网页界面，
输入库存末销数据盘点 / VIN 清单 / 月度目标，输出 **标准 DMS 可导入 Excel** 和离线 HTML 看板。

- 自动分店（按网点权重水位法，每店保底留 1 台）
- 末销日期均匀铺到当月工作日（避免日销峰谷）
- 存销比体检（在库 / 在途 / 合计三套比值，自动给「偏高 / 偏低」增减建议）
- 新 VIN 无需命中快照；车型由 VIN 前缀 + `config/models.json` 双重推断
- 所有逻辑在本机运行，数据不外传

---

## 快速开始

### 1. 数据准备

```
车辆库存_YYYY-MM-DD.xlsx          # 平台导出的全量车辆台账 (31 列)
VIN 清单 (txt/csv/xlsx/xls)         # 入库用；PL 文件名里的项目号会自动抓到「备注」列
月度末销目标.xlsx                   # 看板 / 末销目标用 (Sheet1: 项目=末销, 国家+月份列)
```

放进 `data/inbox/` 就自动识别，也可以在网页里点「上传快照」。

### 2. 双击运行

**直接下载（点一下就下，不用解压）**

| 平台 | 安装包 | 下载链接 |
|------|--------|----------|
| macOS | `vehicle-flow-macos.dmg` | [点此下载](https://github.com/dali-huang/vehicle-flow/releases/latest/download/vehicle-flow-macos.dmg) |
| Windows | `vehicle-flow-windows.exe` | [点此下载](https://github.com/dali-huang/vehicle-flow/releases/latest/download/vehicle-flow-windows.exe) |

也可以直接打开 [Releases 页面](https://github.com/dali-huang/vehicle-flow/releases) 选最新版本。

程序启动后会自动打开 `http://127.0.0.1:8787`（关闭 30 秒后自动退出，不常驻）。

**macOS 首次打开：弹"Apple 无法验证是否包含恶意软件"是正常的**（未做 Apple 公证的第三方 app 都会被 Gatekeeper 拦一次）。打开方法：

- 方式一（推荐）：在「访达 → 应用程序」里找到 `vehicle-flow`，**右键 → 打开**，弹窗里再点「打开」
- 方式二：打开「系统设置 → 隐私与安全性」，拉到底点「仍要打开」
- 方式三（一次性解除拦截）：终端运行
  ```bash
  xattr -dr com.apple.quarantine /Applications/vehicle-flow.app
  ```
  然后正常双击即可

**Windows 首次运行**：如遇 SmartScreen 提示，点「更多信息」→「仍要运行」。

> 说明：本程序是本地运行的离线工具，不联网、不上传数据。若需要彻底免拦截，可用 Apple Developer ID（$99/年）做正式签名+公证；目前为免费 ad-hoc 签名版本。

### 3. 命令行方式

```bash
# 首次使用，自动生成 config/countries/<进口商>.json 主数据
python3 vehicle_flow.py init

# 入库 (新 VIN 本来就不在快照里，正常)
python3 vehicle_flow.py intake \
  --vin-file data/inbox/batch_01.xls \
  --importer IMP001 --date 2026-11-03

# 末销 (70 台，自动铺到当月工作日)
python3 vehicle_flow.py sale --importer IMP001 --month 2026-11 --count 70

# 末销 (按月度目标文件)
python3 vehicle_flow.py sale --importer IMP001 --month 2026-11 --source carflow

# 末销识别 (上传门店真实末销表, 自动核对库存后出标准文件)
python3 vehicle_flow.py sale-file --file data/inbox/末销明细.xlsx

# 看板
python3 dashboard.py
python3 vehicle_flow.py dashboard
```

输出都在 `out/<日期>/` 下面，Excel 文件名自动带批次项目号。

---

## 入库文档字段

标准 8 列（DMS 导入格式）：

| VIN | 车型 | 进口商编码 | 进口商入库时间 | 经销商编码 | 经销商入库时间 | 末销时间 | 备注 |
|-----|------|----------|------------|----------|------------|---------|------|
| 17 位字母数字 | DMS 车型码 | IMP001 | 2026-11-03 | IMP001002 | 2026-11-03 | (末销时填) | PR26AB-001 |

- 新 VIN：车型从 PL 文件名 + `config/models.json` 推断；备注列自动写入文件名里的项目号
- 已在库存：直接沿用快照里的车型与状态
- 所有日期 `YYYY-MM-DD` 文本；同 VIN 不重复提交

---

## 末销逻辑

末销只有一个入口，进去后切换两种方式，输入与处理互不影响：

| 方式 | 输入 | 适用场景 |
|------|------|----------|
| **按数量生成**（sale） | 报台数（我报量） | 给自己定出库量，系统帮你分店、选车、铺日期 |
| **上传真实末销表**（sale-file） | 门店实际末销明细表 | 门店已经出库，把真实结果转成 DMS 可导入格式 |

**按数量生成** 的四步：

1. **报量与参考**：填你要末销的台数；日志显示月均 A（近 6 个月）、目标存销比与可执行上限作参考，不自动替你填数
2. **月均**：近 6 个完整月（最近半年）的平台「末销时间」；窗口零销则回退历史均值 → 月度目标值
3. **分店**：网点权重水位法，每店至少留 `min_keep_units` 台（默认 1）
4. **选车**：同店 FIFO（生产日期 / 创建时间升序）
5. **日期**：当月工作日均匀铺（跳过周末），峰谷差 ≤ 1；未来日期只取 ≤ 今天的那部分

**末销识别** 的处理：

- 表头自适应，不要求固定列名：`VIN / 车架号`、`车型 / 商品名称`、`经销商编码 / 门店`、`末销时间 / 出库日期`、`数量`
- 有 VIN 的行：按快照核对状态（只有在库的能出），重复 VIN、非在库状态、快照里找不到的一律拦下并列出原因
- 只给「经销商 + 数量」的行：自动从该店在库车里按 FIFO 取对应台数，库存不足会提示
- 输出的仍是标准 8 列导入文件，同时写末销台账防重复

## 存销比看板

- **分国家 / 分车型 / 单店** 三个维度的末销与库存
- **在库 / 在途 / 合计** 三套存销比
- **目标带**（`target_ratio ± band`）内为 `ok`，`high` 提示多出量，`low` 提示补货量
- 目标在 `config/targets.json` 按国家设：

```json
{
  "window_months": 6,
  "default": { "target_ratio": 2.0, "band": 0.3 },
  "by_country": {
    "国家B": { "target_ratio": 2.5, "band": 0.4 }
  }
}
```

---

## 全球适用

程序不硬编码区域，任何国家都能跑：

| 配置 | 作用 | 自动生成 |
|------|------|---------|
| `config/countries/<进口商>.json` | 进口商 / 国家名 / 经销商列表 / 网点数 | 首次导入自动从快照生成 |
| `config/models.json` | 车型码 ↔ VIN 前缀映射 | 已有通用模板，可按自家车型表更新 |
| `config/targets.json` | 存销比目标（分国家） | 首次运行自动带默认值 |

要支持一个新市场，只需：快照里的进口商编码不一样 → 程序自动建一个 `<进口商>.json`；车型码不一样 → 改 `models.json` 里的 `prefixes`。

---

## 目录结构

```
vehicle-flow/
├─ config/
│   ├─ countries/*.json       # 各国进口商 + 经销商主数据 (运行时自动生成)
│   ├─ models.json            # 车型 ↔ VIN 前缀
│   └─ targets.json           # 存销比目标 (window_months / default / by_country)
├─ data/
│   ├─ inbox/                 # 库存末销数据盘点 / VIN 清单 / 月度目标
│   └─ ledger/*.json          # 本地台账（防重复）
├─ out/
│   ├─ 末销日期/              # 入库 Excel
│   ├─ YYYY-MM/               # 月度末销 Excel
│   └─ dashboard/             # 看板 HTML
├─ vehicle_flow.py            # 主程序 (intake / sale / dashboard / check)
├─ dashboard.py               # 看板生成
├─ webapp.py                  # 本地网页界面入口
├─ .github/workflows/build.yml # GitHub Actions: 同时产出 exe + dmg
└─ build_windows.bat          # Windows 本地打包
```

---

## 技术栈

- Python 3.9+
- openpyxl (xlsx) + xlrd (xls)
- 本地网页端：标准库 `http.server` + 原生 JS，零依赖
- 打包：PyInstaller（macOS .app → dmg，Windows .exe）

---

## 常见问题

**Q: 新 VIN 不在快照里是不是坏了？**
A: 正常。程序只拦重复提交（台账防重）。新 VIN 的车型从 PL 文件名 / VIN 前缀推断。

**Q: 看板月均销量为什么偏低？**
A: 快照导出有延迟时，最近一个月的末销记录可能不全。看板顶栏会提示「XX 月零销月，A 被低估」。重新导一份完整快照即可。

**Q: SPO 实力值在哪配？**
A: `config/countries/<进口商>.json` 里每个 dealer 加 `"spo": 120`。没配时用网点数作权重。

## 许可证

MIT.
