# TK 菲律宾表格转化工具 v2.0

把 EasyBoss 导出的 `#SKU` 表格（或 TikTok 批量上传模板本身）一键转成
**TikTok Shop 菲律宾站** 的批量上传模板 V5.0.2。

> 这是「TK 菲律宾表格转化工具 v1.2」的完整 Python 重写版（tkinter + openpyxl），
> 跨平台源码 + Windows 一键打包。逻辑、字段映射、输出格式与原版一致。

---

## 功能

| 功能 | 说明 |
| --- | --- |
| 转化 | 源 xlsx → TikTok V5.0.2 模板，自动套用设置（标题前缀/品牌/价格/尺码/包裹等） |
| 入池 | 把一批产品追加到命名产品池（可多次入池） |
| 提取 | 从产品池中读出产品，可直接再次转化 |
| 设置 | 全字段可配置：标题、品牌、类目、价格、库存、COD、标准尺码、描述、尺码图、包裹尺寸、`pre_order_time`、`delivery` |
| 源表识别 | 自动识别 EasyBoss 导出#SKU 或 TikTok 批量上传模板；表头列名可在设置页手动调整 |
| 多副本 | 每产品可生成 N 个 listing（带随机后缀），方便批量铺款 |
| 进度 | 底部真实进度条 + 阶段状态；所有耗时操作走后台线程，界面不假死 |

---

## 目录结构

```
tk-ph-converter/
├── app/
│   ├── __init__.py
│   ├── main.py             # tkinter GUI（入口：main()）
│   ├── config.py           # config.json 读写 / 默认配置
│   ├── source_reader.py    # EasyBoss / TikTok 输入解析 + Product 模型
│   ├── tiktok_writer.py    # 写 TikTok V5.0.2 模板
│   ├── converter.py        # 核心转化流程
│   └── product_pool.py     # 入池 / 提取
├── assets/
│   └── batch-product-source.xlsx   # TikTok V5.0.2 模板（随包分发）
├── TK菲律宾表格转化工具.spec       # PyInstaller 配置
├── build_windows.bat               # Windows 一键打包脚本
├── requirements.txt
├── config.json                     # 用户设置（首次运行后生成）
└── README.md
```

---

## 快速开始（开发模式 · Mac / Windows / Linux 都行）

```bash
# 1) 装依赖（只需 openpyxl，tkinter 随 Python 发行版自带）
python3 -m venv .venv
source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 2) 跑 GUI
python -m app.main
```

---

## 打包成 Windows .exe（推荐给你的客户使用）

在一台 **Windows** 机器上：

1. 安装 **Python 3.11+**（勾选 `tcl/tk and IDLE` 和 `Add python.exe to PATH`）
2. 把整个 `tk-ph-converter/` 文件夹拷到 Windows 上
3. **双击 `build_windows.bat`**

脚本会：
- 创建 `.venv`
- 装依赖 + PyInstaller
- 运行 PyInstaller → `dist\TK菲律宾表格转化工具.exe`
- 复制 `assets\` 到 `dist\`
- 提示构建完成

把 `dist\` 整个目录打成 zip 发给客户即可。客户双击 exe 直接运行，**无需装 Python**。

> ⚠️ 当前 Mac 为 arm64 且无 Wine，无法在 Mac 上交叉编译出 Windows .exe。
> 你的 Windows 虚拟机里跑一下 `build_windows.bat` 即可。

---

## 字段映射（EasyBoss 导出#SKU → TikTok V5.0.2）

| TikTok 列 | 源（EasyBoss） | 说明 |
| --- | --- | --- |
| category | 产品类目 | 可在设置里强制覆盖为指定类目 |
| brand | 品牌 | 同上 |
| product_name | 产品名 | 可加标题前缀 + 随机后缀 |
| product_description | 产品描述 | 同上 |
| main_image, image_2..9 | 产品图片 1..9 | 直接透传 |
| property_name_1 / value_1 | 规格1名称 / 规格1选项 | 多 SKU 去重后逗号拼接 |
| property_1_image | SKU图片 | 取首条变体 |
| property_name_2 / value_2 | 规格2名称 / 规格2选项 | 启用"标准尺码"时改为 S,M,L,XL,2XL,3XL |
| parcel_weight / length / width / height | 包裹重量（KG）/ 尺寸（CM） | KG → g 自动换算；可整体覆盖 |
| price | 税前价格 | 可强制覆盖 |
| quantity | 库存 | 同上 |
| pre_order_time | — | 设置里手动填（默认 3） |
| seller_sku | 平台SKU | 取 master（自动剥掉 `-颜色-尺码` 后缀） |
| size_chart | 尺码图 | 可整体覆盖 |
| cod | — | 设置里手动填（默认 Y） |
| delivery | — | 设置里手动填 |

EasyBoss 是一行一个 SKU，所以工具会**按产品名分组**再合并变体，每个产品在 TikTok
模板里生成 1 行（开启「每产品输出份数」则生成 N 行带不同随机后缀的 listing）。

源表列名映射在「设置 → 源表列名映射」里可改，默认是 EasyBoss 的中文表头；
切换其它 ERP 只需要改这一组映射即可。

---

## 配置文件 `config.json`

首次运行后会在 exe 同目录生成（或源码运行时在项目根目录）。结构：

```jsonc
{
  "version": 2,
  "product_xlsx_last": "...",          // 最近打开的源表
  "product_xlsx_output_dir": "...",    // 最近使用的输出目录
  "product_pool_dir": "...",           // 产品池目录
  "product_xlsx_settings": {
    "title_prefix_enabled": true,
    "title_prefix": "COD Unisex T-shirt【S-3XL】 ",
    "brand_enabled": true, "brand_value": "No brand",
    "price_enabled": true, "price_value": 356,
    "quantity_enabled": true, "quantity_value": 999,
    "cod_enabled": true, "cod_value": "Y",
    "fill_sizes_enabled": true,
    "standard_sizes": "S,M,L,XL,2XL,3XL",
    "title_random_suffix_enabled": true,
    "random_suffix_length": 3,
    "category_enabled": true, "category_value": "Men's Tops/T-shirts",
    "output_copies": 2,
    "description_enabled": true, "description_value": "...",
    "size_chart_enabled": true, "size_chart_value": "...",
    "parcel_enabled": true,
    "parcel_weight_value": 200, "parcel_length_value": 10,
    "parcel_width_value": 10, "parcel_height_value": 5,
    "pre_order_time_value": 3,
    "delivery_value": ""
  },
  "source_column_mapping": { /* 字段 → 源表列名 */ }
}
```

---

## 与原版 v1.2 的差异

| 项 | 原版 v1.2 | 本版 v2.0 |
| --- | --- | --- |
| 平台 | 仅 Windows .exe | 跨平台源码 + Windows 一键打包 |
| 界面 | tkinter | 同（外观一致） |
| 模板 | V5.0.2 | V5.0.2（含 pre_order_time） |
| 产品池 | 同 | 同（向后兼容 `_pools_meta.json`） |
| 进度 | 真进度条 | 同 |
| 第三方依赖 | 内嵌于 exe | 源码模式只需 openpyxl |

---

## 常见问题

**Q: 打包后 exe 启动报错「找不到模板」？**
A: 确认 `dist\assets\batch-product-source.xlsx` 存在。`build_windows.bat` 末
段会从 `assets\` 复制过去；如果是手动打包，漏了这一步就会报错。

**Q: 我的 EasyBoss 导出的列名对不上？**
A: 打开「设置 → 源表列名映射」，把对应字段改成你源表里的实际列名即可。

**Q: output_copies 设为 2 会不会被 TikTok 拒？**
A: 同一产品生成 N 份 listing 是常见做法（不同标题便于赛马），但请遵守
TikTok Shop 的发布规则。

**Q: description / size_chart 想用源表的，不要覆盖？**
A: 在「设置」里把对应 `*_enabled` 勾掉即可。
