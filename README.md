# <div align='center' >Auto Check</div>

<div align='center'>A tool for quick check and marking at Microsoft Office files(docx, xls, ppt etc.)</div>


<div align=center><img width="500" src="/pictures/gui_overview.png"/></div>



<p align="center">
<a href="https://github.com/tlj-ryuu"><img src="https://img.shields.io/badge/Author-TLJ-purple.svg"></a>
<a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.9-14354C.svg?logo=python&logoColor=white"></a>
<a href="https://www.microsoft.com/zh-cn/software-download/windows10%20"><img src="https://img.shields.io/badge/Platform-windows10-green.svg"></a>
<a href="https://www.microsoftstore.com.cn/software/office"><img src="https://img.shields.io/badge/Microsoft_Office-docx_|_xls_|_xlsx_|_ppt_|_pptx-orange.svg"></a>
<a href="https://raw.githubusercontent.com/onevcat/Kingfisher/master/LICENSE"><img src="https://img.shields.io/badge/License-MIT-aqua.svg"></a>
</p>

# 1. Introduction

AutoCheck, in general, is the initial product of a project development, is a link in the complete project. The assumed application scenario of the project is to open a file of a specified category for review according to the input signal (which can be voice input, electronic signal, any signal that can distinguish categories), and then mark the file by giving a signal (which can be voice signal, keyboard signal, etc.) through manual judgment. In the future, it is hoped that it can be extended to give classification signals after natural language understanding through artificial voice input, and the signal that gives whether to mark can also be connected to the natural language understanding module. Whether to make a mark can also be changed from a manual judgment to an automatic judgment to achieve a fully automatic effect.


For the above purpose, as an initial product, an application with gui interface is implemented, and the numbers obtained by gui are used as signals. Meet the following requirements:


- [x] Automatically open a disk file under win according to the signal: input numbers in the gui, open the specified format file, able to change the file path where files are opened

- [x] Opening a file and entering the second file can automatically close the first file, ensuring that the interface has only one file, preventing forgetting to close and causing a large number of files to be opened

- [x] The content of the file can be marked (such as word changing from white background to red) after the open file is artificially given a signal. The artificial signal is temporarily realized by checking the selection box on the gui

- [x] Package as a complete application that can be used by people who don't have a python interpreter, i.e., package the entire environment inside.

- - -
本工具，总体上来看是一个项目开发的初步产品，是完整项目的一个环节。项目假设的应用场景是根据输入信号（可以是语音输入，电子信号，任何能区分类别的信号）打开指定类别的文件进行审阅，再通过人工判断给出信号（可以是语音信号，键盘信号等）对文件进行标记。未来希望能扩展成通过人工语音输入，进行自然语言理解后给出分类信号，同样给出是否做标记的信号也可以接入自然语言理解模块。对于是否做标记也可以从人工判断改为自动判断，以达到全自动效果。

为此，初步实现了一个带有gui界面的应用，以gui获得的数字作为信号传递。满足以下需求：

1. 根据信号自动打开win下某盘文件：在gui输入数字，打开指定格式文件，能改指定位置

2. 打开一个文件后输入第二个文件能自动关闭第一个文件，保证界面只有一个文件，防止忘记关闭造成大量文件被打开

3. 对打开的文件人工给定信号后能对文件内容进行标记（比如word从白色背景改成红色）人工信号暂时以gui上选择框进行勾选来实现

4. 打包成完整的应用程序，对于没有python解释器的人也能用，即把整个环境打包进去。

# 2. Applied range

* Microsoft Office Products (especially not for WPS)
* \> Windows10

# 3. Getting Started

## 3.1 Prerequisites
* \> Python3.9
  
*My development environment is Python3.9 and I have not tested a lower version so I recommend using the same version of Python*

## 3.2 Usage from code

After git this repo, just run the main program [**AutoCheck_back.py**](/AtuoCheck_back.py) in the python programming environment

```
$ python AutoCheck_back.py
```

## 3.3 Usage from application

AutoCheck tool is packaged as an exe application by Pyinstaller. 3 versions of exe can be found under the [EXEs](/EXEs) folder:

| Name | Description |
| --- | --- |
| AutoCheck.exe | One file exe |
| AutoCheck_cs.exe | One file with console |
| AutoCheck_ofcs | One folder with console |

:mushroom: Since the print output statement in the background could cause bugs on some computers with no programming environment at all, two additional versions of the exe with console were made

### 3.3.1 Steps

**step1: Set file path and push "Set" button**

<div align=center><img width="530" src="/pictures/setpath.png"/></div>
<div align=center>pic1: file path</div>

<div align=center><img width="530" src="/pictures/fileset.png"/></div>
<div align=center>pic2: push "set" button</div>

**step2: Choose signal and push "Check" button**

<div align=center><img width="530" src="/pictures/step2.png"/></div>
<div align=center>pic3: push "check" button</div>

**step3: If decided to mark push "Confirm" button**

<div align=center><img width="530" src="/pictures/step3.png"/></div>
<div align=center>pic4: background changed to red</div>

**step4: return to step2 if wanna check next file**

<div align=center><img width="530" src="/pictures/step4_1.png"/></div>
<div align=center>pic5: open excel according to signal</div>

<div align=center><img width="530" src="/pictures/step4_2.png"/></div>
<div align=center>pic6: mark excel</div>

<div align=center><img width="530" src="/pictures/step4_3.png"/></div>
<div align=center>pic7: open ppt according to signal</div>

<div align=center><img width="530" src="/pictures/step4_4.png"/></div>
<div align=center>pic8: mark ppt</div>

### 3.3.2 Notice

Some Settings and prompt statements to prevent violation operation which could lead to a fatal error

Situations and correspoding are as below:

* when push check button first without setting file path
<div align=center><img width="530" src="/pictures/caution1.png"/></div>

* when push confirm button first without checking first
<div align=center><img width="530" src="/pictures/caution2.png"/></div>

* when set wrong path
<div align=center><img width="530" src="/pictures/caution3.png"/></div>

* when check non-exsiting type file
<div align=center><img width="530" src="/pictures/caution4.png"/></div>

## 3.4 Setting

放个信号对应的表格，以及对应的ext,以及对应的标记

## 3.5 Demo
放个视频
# 4. Bugs
win7 试试使用-win-private打包方式
# 5. Contributors
<a href="https://github.com/tlj-ryuu/AutoCheck/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=tlj-ryuu/Autocheck" />
</a>
