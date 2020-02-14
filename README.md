# 人名空格补齐实用程序 excel-name-space-formatter (nsf) <!-- omit in toc -->

本程序可以快速批量对 Excel 中指定单元格内由两个汉字组成的姓名中间添加空格，使其*看起来*与三个汉字组成的姓名“两端对齐”。

## TOC <!-- omit in toc -->

- [下载安装](#%e4%b8%8b%e8%bd%bd%e5%ae%89%e8%a3%85)
  - [GitHub Release](#github-release)
  - [源代码](#%e6%ba%90%e4%bb%a3%e7%a0%81)
- [从源码构建EXE](#%e4%bb%8e%e6%ba%90%e7%a0%81%e6%9e%84%e5%bb%baexe)
  - [安装依赖](#%e5%ae%89%e8%a3%85%e4%be%9d%e8%b5%96)
  - [构建可执行文件](#%e6%9e%84%e5%bb%ba%e5%8f%af%e6%89%a7%e8%a1%8c%e6%96%87%e4%bb%b6)
  - [清理](#%e6%b8%85%e7%90%86)

## 下载安装

使用前请确保已经安装了 Excel。

### GitHub Release

二进制文件~~绿色无毒~~，下完即可运行，干净无依赖，不需要安装 Python。

[GitHub Release](https://github.com/Z4HD/excel-name-space-formatter/releases)

### 源代码

```shell
git clone --depth=1 https://github.com/Z4HD/excel-name-space-formatter.git
```

安装依赖后即可运行

```shell
pipenv install
pipenv run main
```

## 从源码构建EXE

首先确保已经安装 `pipenv`

```shell
pip3 install pipenv
```

### 安装依赖

进入源代码目录后执行

```shell
pipenv install --dev
```

### 构建可执行文件

```shell
pipenv run build
```

可以在 `dist\` 文件夹中找到构建完成的EXE。

### 清理

删除构建过程中产生的临时文件，**不包括** `dist\` 目录下的可执行文件。

```shell
pipenv run clean
```

删除构建过程中产生的所有临时文件，**包括 `dist\` 目录下的可执行文件。**

```shell
pipenv run cleanall
```
