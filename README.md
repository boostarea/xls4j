# xls4j
[![Build Status](https://travis-ci.org/ctripcorp/apollo.svg?branch=master)](https://travis-ci.org/ctripcorp/apollo)
[![GitHub release](https://img.shields.io/github/release/ctripcorp/apollo.svg)](https://github.com/ctripcorp/apollo/releases)[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)



Xls4j (大批量数据处理) 能处理紧急数据处理需求. 能够根据excel文件,生成大批量的指定sql语句, 并且具备拆分大型excel文件的能力.

基于Spring Boot和Apache poi开发,打包后可直接运行,不需额外安装应用容器.

演示环境 (Demo):

-[129.204.32.196:8080](http://129.204.32.196:8080/)

# Screenshots

![运行页面](https://raw.githubusercontent.com/wiki/ooooor/siren/screenshot.png)

# Features
* **自定义SQL语句,根据excel大批量生成**
  * 在apache poi之上, 针对io流进行封装,解析/写入excel速度明显提升。

* **快速生成结果(高并发) **
  * 线程间通讯,(v1.0.0使用Redis) v1.0.1后,将全面使用Spring自带Cache的ConcurrentMap进行管理。
  * 实现框架的对外部依赖尽可能少,部署非常简单,只需jar包即可运行.

* **历史回溯功能**
  * 可以方便的找回已生成的文件(默认数量:50)。

* **支持大文件的拆分**

* **支持Docker镜像部署*

* **源码完全开源**

# Usage
  ```text
nohup java -jar /data/projects/xls4j/exl4j-1.0-SNAPSHOT.jar > /data/log/xls4j/xls4j.log 2>&1 &
```

# Design

# Development

# FAQ

# Support 

# Contribution
  * Source Code: https://github.com/ooooor/xls4j
  * Issue Tracker: https://github.com/ooooor/xls4j/issues