### 功能描述：

    将深圳市电子发票贴到公司报销单模板中，并导出为pdf文件。

### 注意事项：

只能运行在windows上，且已经安装了ms office。

### 依赖：

- python 3.7
- python-pptx 负责操作pptx
- PyMuPDF 负责操作pdf
- comtypes 负责调用系统已安装的ppt程序，导出pdf
- PyPDF2用于合并多个pdf文件

### 安装：

1. 安装python3.7

2. 安装对应的第三方库

```
pip install -r requirements.txt
```

### 使用步骤：

1、将电子发票pdf文件放到"inputs"文件夹；

2、运行如下脚本，在"outputs"文件夹生成的pdf票据文件；

```shell
python .\invoice_converter.py
```

### 实现思路：

完全按照手动贴票的流程进行：

电子发票pdf------>PNG图片------>插入pptx模板（调整位置、大小）------>导出pdf文件

