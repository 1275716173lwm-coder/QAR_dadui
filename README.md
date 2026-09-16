# QAR批量分析程序

## 环境

- Windows 10/11
- Python 3.11或更高版本

## 安装与运行

在本目录打开PowerShell，执行：

```powershell
python -m pip install -r requirements.txt
python qar_app.py
```

启动后选择包含人员子文件夹的数据目录，点击“开始分析”。分析过程中可以安全取消。

## 生成Windows EXE

打包操作只需要在Windows上执行一次。建议先创建独立虚拟环境，避免本机其他Python包进入EXE：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
```

在项目根目录执行以下命令：

```powershell
pyinstaller --noconfirm --clean --onefile --windowed --name "QAR批量分析" --add-data "result_sample;result_sample" qar_app.py
```

打包完成后，可执行文件位于：

```text
dist\QAR批量分析.exe
```

可以将这个EXE复制到其他Windows电脑直接运行，不需要安装Python。`result_sample`模板已经包含在EXE中；运行结果会写入EXE所在目录下的`result`文件夹。

如果程序在目标电脑无法启动，可暂时去掉`--windowed`重新打包以显示错误信息：

```powershell
pyinstaller --noconfirm --clean --onefile --name "QAR批量分析_调试" --add-data "result_sample;result_sample" qar_app.py
```

重新打包前可以删除旧的`build`、`dist`和`QAR批量分析.spec`；这些都是PyInstaller生成物，不影响源代码。

## 输出

- 结果表：`result/qar_YYYYMMDD_HHmmss.xlsx`（使用保存完成时间）
- 分析报告：`result/qar_report_YYYYMMDD_HHmmss.docx`
- 日志：`result/qar_log_YYYYMMDD_HHmmss.txt`

每次成功分析都会从`result_sample/result_sample.xlsx`重新生成结果表。Excel、DOCX和日志均不会覆盖历史文件；同一秒内重名时自动追加`_1`、`_2`等序号。日志采用UTF-8编码和制表符分隔。

## 数据目录要求

所选目录的直接子文件夹名称作为人员姓名。每个人员目录中的CSV只有在文件名恰好包含一种以下标记时才会分析：

- `YZDCFM`
- `YZDLEAP`
- `YZDPW`

D、E列统计50ft AGL至接地期间舵量绝对值超过15的采样点数及其占全部有效舵量采样点的比例。L列统计疑似接地前反向蹬舵航班数：接地交叉角绝对值严格大于3°，且50ft至接地期间存在与交叉方向一致的非零舵量时，每个航班最多计1次；该判断不要求舵量绝对值大于15。接地交叉角严格大于20°时视为错误数据，不参与接地交叉角和反向蹬舵判断；恰好20°仍按正常数据处理。

DOCX会列出接地交叉角超限事件的侧风分量、抬头速率超限事件的最大抬头速率和高高原判断结果，以及疑似接地前反向蹬舵事件的航班信息、带方向交叉角和代表舵量证据。
