# Excel 指定列生成条形码工具

一个 Windows 桌面小工具：读取 Excel 文件，按指定列把每个值批量生成一维条形码
（Code128 / Code39 / EAN13 / EAN8 / UPCA）并导出为 PNG 图片。

## 功能

- 选择 Excel 文件，**自动读取列名**（支持自定义"列名行"，默认第 2 行）
- 自定义"数据起始行"（默认第 3 行，自动跟随列名行 +1）
- 选择条形码类型（code128 / code39 / ean13 / ean8 / upca）、是否显示文字
- 选择导出目录，一键批量生成
- 内置自检模式，便于在打包后的 exe 客户端验证生成能力：

  ```bat
  BarcodeTool.exe --selftest <excel路径> [输出目录] [列名行]
  ```

## 运行方式

- **免安装（已发布）**：直接下载仓库里的可执行文件
  [`BarcodeTool.exe`](https://github.com/Wflysea/table_to_barcode/raw/main/dist/BarcodeTool.exe)（约 19MB，PyInstaller 单文件打包，双击即用）
- **源码运行**：`python app.py`（需安装 `openpyxl`、`python-barcode`、`Pillow`，且 Python 自带 `tkinter`）

## 下载

| 文件 | 说明 | 链接 |
| --- | --- | --- |
| `BarcodeTool.exe` | 已编译的 Windows 单文件程序 | [点此下载](https://github.com/Wflysea/table_to_barcode/raw/main/dist/BarcodeTool.exe) |

> 该 exe 已随仓库 `main` 分支提交（绕过 `.gitignore` 的 `dist/` 忽略规则），可随时从上面链接获取最新版。

## 注意事项

一维条形码（Code128 等）只能编码 ASCII（数字 + 英文），**中文列会被跳过**。
例如"物品名称"列是中文无法生成，应改用数字编码列（如"物品编码"）。

## 文件结构

| 文件 | 说明 |
| --- | --- |
| `app.py` | tkinter 图形界面主程序 |
| `barcode_tools.py` | 条形码生成核心逻辑（读取列名 / 生成条码 / 字体回退） |
| `test_flow.py` | 功能自测脚本（含两种列名行场景） |
| `run.bat` | Windows 启动脚本 |

## 打包为 exe

```bat
pyinstaller --onefile --windowed --name BarcodeTool ^
  --hidden-import=openpyxl --hidden-import=barcode --hidden-import=PIL app.py
```

> 打包后若显示文字异常，请确保程序能在运行时找到系统字体（Windows 默认
> `C:\Windows\Fonts\arial.ttf` 等），否则会自动关闭文字以保证条码本体可生成。
