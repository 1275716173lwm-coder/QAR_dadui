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

## 输出

- 结果表：`result/qar.xlsx`
- 日志：`result/qar_log_YYYYMMDD_HHmmss.txt`

每次成功分析都会从`result_sample/result_sample.xlsx`重新生成结果表。日志采用UTF-8编码和制表符分隔，不覆盖历史日志。

## 数据目录要求

所选目录的直接子文件夹名称作为人员姓名。每个人员目录中的CSV只有在文件名恰好包含一种以下标记时才会分析：

- `YZDCFM`
- `YZDLEAP`
- `YZDPW`

D、E列当前保持空白；侧风和文档分析当前不执行。
