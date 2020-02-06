# excel-name-space-formatter (nsf)

A people name formatter in M$ Excel which can add or remove space in the center of a name. ONLY CHINESE USER NEED IT.

## 下载安装

### GitHub Release

二进制文件~~绿色无毒~~，下完即可运行，干净无依赖，不需要安装 Python。

Comming soon...

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
