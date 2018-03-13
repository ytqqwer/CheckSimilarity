# CheckSimilarity 1.0

<!-- TOC depthFrom:2 -->

- [1. 软件简介](#1-软件简介)
    - [1.1. 文档说明](#11-文档说明)
    - [1.2. 注意事项](#12-注意事项)
    - [1.3. 系统要求](#13-系统要求)
    - [1.4. 安装说明](#14-安装说明)
- [2. 界面说明](#2-界面说明)
    - [2.1. 菜单栏-文件](#21-菜单栏-文件)
        - [打开](#打开)
        - [保存](#保存)
        - [退出](#退出)
    - [2.2. 菜单栏-帮助](#22-菜单栏-帮助)
        - [关于](#关于)
    - [2.3. 主界面](#23-主界面)
        - [搜索](#搜索)
        - [选择词类](#选择词类)
        - [GKB词语列表](#gkb词语列表)
        - [现汉词语列表](#现汉词语列表)
        - [Check](#check)
        - [上一个](#上一个)
        - [下一个](#下一个)
- [3. 操作流程](#3-操作流程)
- [4. 其他](#4-其他)
    - [4.1 使用提示](#41-使用提示)
    - [4.2 常见问题](#42-常见问题)

<!-- /TOC -->

## 1. 软件简介
显示同形词语间的对应关系，支持手动选择相似度最高的词语，并记录处理结果。

### 1.1. 文档说明
本文档是『CheckSimilarity』（以下简称“软件”）的配套说明文档。

本文档使用简洁易用的[Markdown](https://www.appinn.com/markdown/)标记语言编写。

推荐使用能够解析Markdown语言的编辑器[Visual Studio Code](https://code.visualstudio.com/)。

### 1.2. 注意事项
软件只支持读取xlsx的Excel文件，并且对列名有严格的要求。文件格式与列名要求可以参考程序附带的测试数据文件。确保所有需要读取的文件都存放在同一目录中。

软件使用正则表达式检索指定目录下的xlsx格式文件，文件命名规则参考测试数据文件命名方法。其中代表映射关系的“to”字段以及代表词类的单字字段必须存在，并且按照从左至右的顺序出现，中间存在其他分割字符，其他字段不做要求。

映射关系参考“to”两端的字符，如果是数字，就被识别成为1；如果是字母，就被识别成m或m。

### 1.3. 系统要求
在使用本软件之前，系统配置需符合以下条件：

- 操作系统要求：Windows 7/8/8.1/10 (32位/64位) ;
- 内存要求：不少于128M;
- 磁盘空间要求：不少于50M剩余磁盘空间；

### 1.4. 安装说明
将压缩包解压至任意目录即可。

软件运行时需要使用 Microsoft Visual C++ Redistributable 2015 运行组件，以运行使用Visual C++编写的应用程序。

相关运行组件已经附带在安装包里，安装2015版即可使用。如不能运行，尝试安装较低版本vcredist。

注意操作系统是64位还是32位。对于32位的操作系统，只能安装32位的运行组件，对于64位的操作系统，除了需要安装64位的组件，32位的运行组件也要安装一遍。
 
## 2. 界面说明

### 2.1. 菜单栏-文件

#### 打开
选择xlsx格式文件所在目录。

#### 保存
将当前的Check结果写入记录文件中。

#### 退出
退出应用程序，同时自动将Check结果写入记录文件。

### 2.2. 菜单栏-帮助

#### 关于
软件的版本说明，作者的联系邮箱。

### 2.3. 主界面
在软件尚未读取到任何文件时，主界面上的控件大多被禁用。

#### 搜索
在当前选择的词类下搜索指定词语，输入词语后点击搜索按钮。在编辑框中支持回车操作。

#### 选择词类
只有读取文件后才会产生效果。使用"词类缩写"作为"词类标识"。选择词类后将会刷新主界面。

#### GKB词语列表
显示GKB词语列表中的一个同形词语。如果列中无法显示全部内容，可以拖动列名来拉伸列。

#### 现汉词语列表
显示当前GKB词语对应的数个词语。如果列中无法显示全部内容，可以拖动列名来拉伸列。

在每一行开头可以对词语打勾，但是如果要记录结果必须点击Check。

#### Check
将当前现汉词语列表中的打勾情况记录下来。此时并不会将Check结果写入文件。

#### 上一个
浏览上一个GKB词语。

#### 下一个
浏览下一个GKB词语。

## 3. 操作流程

- 文件-打开-选择目录
- 选择词类
- 通过搜索，『上一个』或『下一个』按钮进行浏览
- 勾选现汉词语，点击Check
- 点击保存，将Check结果写入文件

## 4. 其他

### 4.1 使用提示
- 时常手动保存
- 使用搜索功能回到上一次的浏览地点
- 与软件同目录的『record.xlsx』文件可以手动编辑

### 4.2 常见问题
- 如遇问题，重启软件