# IDF材料提取工具

这是一个用于 AVEVA E3D / Isodraft Isogen IDF 文件的材料提取工具。现在同时提供桌面窗口和本地网页两种入口。

## 功能

- 自动读取 IDF 中的 `-20/-21` 材料编码和材料描述。
- 根据组件记录第 9 个字段的材料指针，统计材料数量。
- 输出材料编码、描述、数量、管径、管道规格、组件类型和唯一组件号。
- 新建 Excel 模式：可一次选择多个 IDF 文件，生成 `管件明细`、`材料汇总`、`文件信息` 三个工作表；`管件明细` 中每个组件单独占一行，首列为 `来源文件名`，并单列输出 `公称管径`、`壁厚等级`、`压力等级`。按 IDF Word 11 的 BOM 控制百位过滤不进入 BOM 的组件；材料汇总按材料、管径、壁厚、压力等级和管道规格分别成行。
- 单位规则：管子/定长管段/变长管段等管道类组件按“米”计量，数量为起点和终点坐标计算出的长度（米）；法兰、阀门、垫片等其他零件按“个”计量。
- 表格首列来源文件名之后增加“管线号”，该字段读取 IDF 头部的 `-6` 记录，批量处理时每个文件分别保留自己的管线号。
- Olet 辅助记录过滤：组件类型为 `40` 或 `42` 且材料描述包含 `PIPE` 时，视为 Olet 几何辅助记录，不重复计入管子材料；正常管子类型 `100/101/102/103/104` 保留。
- 自定义模板模式：选择已有 `.xlsx`/`.xlsm`，自动识别常见中文或英文表头，并从表头后追加“每个组件一行”的明细；如果模板中已有 `管件明细` 或旧版 `组件明细` 工作表，也会同步填充。
- 支持命令行批处理。

## 本地网页工具

先安装依赖：

```powershell
python -m pip install -r .\requirements.txt
```

然后在 PowerShell 中执行：

```powershell
.\start_local.bat start
```

看到下面的提示后，打开 Chrome 或 Edge，访问 `http://127.0.0.1:8765`：

```text
IDF材料网页工具已启动：http://127.0.0.1:8765
```

启动命令会在当前窗口运行网页服务，请保持该窗口打开。关闭窗口或按 `Ctrl+C` 会停止服务。

不要直接双击 `web\index.html`，直接打开 HTML 文件时没有后端接口，上传和 Excel 导出不会工作。请使用 `start_local.bat` 启动。

本地服务开关：

```powershell
.\start_local.bat start   # 启动服务
.\start_local.bat status  # 查看运行状态
.\start_local.bat stop    # 停止服务
```

`status` 和 `stop` 可在另一个 PowerShell 窗口中执行。

服务只监听 `127.0.0.1`，文件不会上传到外部服务器。

## 图形界面

```powershell
python .\idf_material_extractor.py
```

选择 IDF 文件后，可以选择：

1. `新建Excel`：程序自动生成标准结果表。
2. `写入自定义模板`：选择已有 Excel 模板、目标工作表和表头行。表头行留空时自动识别。

自定义模板建议使用下列任意中英文表头：

`材料编码` / `Item Code`、`材料描述` / `Description`、`数量` / `Qty`、`单位` / `Unit`、`管道规格` / `Spec`、`公称管径` / `Bore`、`唯一组件号` / `Unique ID`。

输出中的“管径/尺寸字段”对应 IDF 第 8 字段。对管子和管件通常是公称管径，对螺栓等非管道组件可能是该组件的尺寸字段。

如果模板缺少某个字段，程序会在表格最右侧补充该列，不会覆盖原有列。

## 命令行

新建结果表：

```powershell
python .\idf_material_extractor.py `
  --idf "D:\path\line.idf" `
  --mode new `
  --output "D:\path\line_材料清单.xlsx"
```

写入自定义模板：

```powershell
python .\idf_material_extractor.py `
  --idf "D:\path\line.idf" `
  --mode custom `
  --template "D:\path\模板.xlsx" `
  --sheet-name "材料表" `
  --header-row 3 `
  --output "D:\path\line_填充结果.xlsx"
```

`--header-row 0` 表示自动识别表头行。

## 常见问题

- 如果提示“`py` 不是内部或外部命令”，不要使用 `py`，改用 `python`。
- 如果提示“`python` 不是内部或外部命令”，请安装 Python 3.11 或更高版本，并在安装时勾选“Add Python to PATH”，然后重新打开 PowerShell。
- 如果启动命令后没有返回新的 `PS>` 提示符，这是正常的，表示网页服务正在前台运行。保持该窗口打开，使用浏览器访问 `http://127.0.0.1:8765`。
- 如果端口被占用，可以先执行 `.\start_local.bat stop`，或关闭占用 8765 端口的程序后重试。

## 打包为可执行文件

如果需要在没有 Python 的电脑上运行：

```powershell
python -m pip install pyinstaller
pyinstaller --noconsole --onefile --name IDF材料提取工具 .\idf_material_extractor.py
```

生成的 exe 位于 `dist` 文件夹。
