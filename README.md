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

启动后选择包含人员子文件夹的数据目录及一个飞机数据库 `.xlsx` 文件，再点击“开始分析”。分析过程中可以安全取消。

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

可以将这个EXE复制到其他Windows电脑直接运行，不需要安装Python。结果表固定读取 `G:\OneDrive - cqu.edu.cn\飞行资料\CA\飞行部\QAR\result_sample\result_sample.xlsx`；运行结果会写入EXE所在目录下的`result`文件夹。

如果程序在目标电脑无法启动，可暂时去掉`--windowed`重新打包以显示错误信息：

```powershell
pyinstaller --noconfirm --clean --onefile --name "QAR批量分析_调试" --add-data "result_sample;result_sample" qar_app.py
```

重新打包前可以删除旧的`build`、`dist`和`QAR批量分析.spec`；这些都是PyInstaller生成物，不影响源代码。

## 输出

- 结果表：`result/qar_YYYYMMDD_HHmmss.xlsx`（使用保存完成时间）
- 分析报告：`result/qar_report_YYYYMMDD_HHmmss.docx`
- 日志：`result/qar_log_YYYYMMDD_HHmmss.txt`

每次成功分析都会从固定模板重新生成13列结果表（姓名、机型分类及原有11项指标）。Excel、DOCX和日志均不会覆盖历史文件；同一秒内重名时自动追加`_1`、`_2`等序号。日志采用UTF-8编码和制表符分隔。

结果表的表头及全部数据单元格均采用水平、垂直居中对齐；同一人员存在多个机型分类时，合并后的姓名单元格也保持居中。

## 数据目录要求

所选目录的直接子文件夹名称作为人员姓名。每个人员目录中的CSV只有在文件名恰好包含一种以下标记时才会分析：

- `YZDCFM`
- `YZDLEAP`
- `YZDPW`

## 飞机数据库与预检

飞机数据库必须是 `.xlsx` 文件，且必须包含 `Registration` 和 `A/C TYPE` 两列。程序忽略数据库注册号的大小写和首尾空格。CSV文件名中首个符合 `B-[A-Z0-9]{4}` 的内容作为注册号，并统一转为大写匹配。

在计算指标前，程序会一次性预检全部候选CSV。数据库无法读取、缺少必要列、重复注册号对应不同机型、CSV缺少注册号或注册号未匹配，都会停止整次分析；此时不生成Excel和DOCX，但仍生成日志并列出全部问题文件。不支持的数据库机型或数据库发动机后缀与文件名YZD标记冲突，只跳过对应文件。

支持的具体机型及输出分类如下：

- `A319CFM`、`A319LEAP` → `A319`
- `A320CFM`、`A321CFM`、`A320PW`、`A320LEAP` → `A320、A321`
- `A321LEAP`、`A321PW` → `A321NEO`

指标按“人员＋机型分类”独立汇总。同一人员只输出实际存在有效文件的分类，顺序为 `A319`、`A320、A321`、`A321NEO`；多行时Excel合并姓名单元格。DOCX采用“人员一级标题 → 机型分类二级标题 → 指标汇总及事件三级标题”的可折叠结构。

D、E列统计50ft AGL至接地期间舵量绝对值超过15的采样点数及其占全部有效舵量采样点的比例。L列统计疑似接地前反向蹬舵航班数：接地交叉角绝对值严格大于3°，且50ft至接地期间存在与交叉方向一致、绝对值严格大于1的人工舵量时，每个航班最多计1次；绝对值小于或等于1视为未进行人工操纵。跑道磁航向为0，或跑道磁航向列全部有效数值始终不变时，判定为数据错误，接地交叉角和反向蹬舵均不进行判断并分别写入日志。接地交叉角严格大于20°时作为排除数据仅写入日志，不参与接地交叉角和反向蹬舵判断，也不累计失败次数；恰好20°仍按正常数据处理。

DOCX会列出接地交叉角超限事件的侧风分量、抬头速率超限事件的最大抬头速率和高高原判断结果，以及疑似接地前反向蹬舵事件的航班信息、带方向交叉角和代表舵量证据。

三类事件中的源CSV文件名均为指向原始CSV绝对路径的本地文件超链接，可在Word中点击打开源文件。

程序兼容首行以“dataframe info”开头的译码CSV；此时自动使用第2行作为字段名。交叉角严格大于20°以及其他“不适用”状态只写入日志，不累计为失败；比例没有有效数据时统一填写0%。
