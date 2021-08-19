# COCTRPG_replay_pptx_generator

COC跑团replay辅助ppt生成工具

[![standard-readme compliant](https://img.shields.io/badge/readme%20style-standard-brightgreen.svg?style=flat-square)](https://github.com/RichardLitt/standard-readme)

读取跑团记录（txt）以及制作好的ppt模板（pptx)，自动读取发言人与对应模板中页面，批量填写对话内容。

功能：

1. 读取[QQ跑团记录着色器](https://logpainter.kokona.tech/)生成或[朗读女](http://www.443w.com/tts/)适配的跑团记录(txt格式)。
2. 读取已经填写的角色列表 或 自动生成新的角色列表(json) 分类玩家身份(KP, PC, 骰子, NPC)
3. 通过用户设置的关键词列表识别自动更换差分立绘（不稳定）
4. 自动生成朗读女适配的跑团记录，包含骰子音效插入，自动识别成功等级（UI未适配）


## 内容列表

- [安装](#安装)
	- [打包版下载](#打包版下载)
	- [开源代码(本仓库内容)](#开源代码(本仓库内容))
- [使用说明](#使用说明)
- [打包方法](#打包方法)
- [相关仓库](#相关仓库)
- [使用许可](#使用许可)


## 安装

### 打包版下载
[百度云](https://pan.baidu.com/s/1HNHWGwti2OlGXFhdEnVDSQ)  提取码: cfo9

### 开源代码(本仓库内容)
这个项目使用 [python3](https://www.python.org/downloads)。

> 包含pyinstaller作为exe-builder。

```sh
$ pip install -r requirements.txt
```

## 使用说明

运行run.py开启UI界面。

```sh
$ python3 run.py
```

具体使用方法见[视频](https://www.bilibili.com/video/BV19B4y1A7nG)。

### 打包方法

> 注意当前directory需在本项目根文件夹里（COCTRPG_REPLAY_PPTX_GENERATOR/）

```sh
$ pyinstaller -F -w -i icon.ico run.spec
```

> 如果有大佬帮忙看一下如何更精简打包文件十分感谢

## 相关仓库

- [python-pptx](https://github.com/scanny/python-pptx) — 基于python的pptx文件生成包。
- [PyQt5](https://github.com/PyQt5) — UI创建。
- [imagehash](https://github.com/JohannesBuchner/imagehash)  — 用于图像对比，寻找立绘所在位置。
- [pyinstaller](https://github.com/pyinstaller/pyinstaller)。


## 使用许可

[MIT](LICENSE) © Charlotte Wei
