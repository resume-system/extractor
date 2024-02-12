# extractor思路

## 简历格式转换

### 文件转换：word转为pdf

发现使用PDF转换后提取信息的效果更好一些，因此第一步需要把word文件转换为pdf文件。这里使用 python 的 win32com 包来实现转换，需要说明一点的是这个包需要调用 windows 下的 word 程序，因此只支持 windows 平台。（PS：也许有其他支持跨平台的依赖库可以使用）

#### 安装依赖库

```shell
pip install pywin32
```

## 简历提取

### 提取思路

基本的提取思路是先把PDF文件的文本内容提取出来，然后通过正则表达式去匹配值。**这就导致了一个问题，姓名只能猜测而无法准确获取。**PDF文本提取使用的是 pdfplumber 这个库，通过以下命令安装：

```shell
pip install pdfplumber
```

## 使用说明

首先，你的简历文件结构应该如下：

```
data
 - 目录一
   - 一些 pdf 或者 word 文件
 - 目录...
   - ...
```

使用时可以直接通过以下方式调用：

```python
python extractor.py data result.xlsx
```

其中 data 代表简历存放的根目录，result.xlsx 代表保存文件名，这两个参数都是可选的，不加则代表使用默认值 data 和resume-data.xlsx

> PS：如果连续提取，会出现多个resume-data.xlsx按01、02排列

## 善后

最后规则提取出来的简历再构建用户本人的知识图谱，所有用户的信息都存在一个图数据库里
