# EDragonKMS-Actbat(ghpymKMS)
### Ps:I`m not good at English sorry！
# E_Dragon的KMS激活脚本（果核剥壳KMS）
***
## Downloads 下载
[V3.0 Download link|V3.0下载链接](https://github.com/EDragon007/EDragonKMS-Actbat/releases/tag/3.0)
[V3.0b-Sysprep Download link|V3.0b部署版下载链接](https://github.com/EDragon007/EDragonKMS-Actbat/releases/tag/3.0b-Sysprep)
***
## Overvier 概述
### Activate Windows&Office conveniently without virus
### 以一种没有病毒的方式轻松激活Windows&Office
#### KMS activation is valid for 180 days, every 7 days will automatically try to activate to extend the validity
#### KMS激活有效期为180天，每7天会自动尝试联网激活延长有效期
***
# README 中文版
## 一、原理
### 0.准备操作
通过`setlocal EnableDelayedExpansion`设置开启变量延迟
通过`mshta vbscript`指令运行一个VB脚本提示符以方便下面指令的运行
通过在`%TEMP%`目录下创建脚本并执行以获取管理员权限，执行后删除
通过`set`指令设置KMS激活服务器以便后续程序统一调用
通过`ping`指令探测`%KMSSERVERR%`的网络是否与客户端直接能正常访问
### 1.Windows的激活
#### 1)检测Windows版本
通过`ver`指令输出命令提示符的版本以识别Windows版本
使用管道符与`find`配合对输出版本号进行识别并使用`goto`指令跳转至相应的激活选项
#### 2)选择合适的KMS密钥
[Microsoft关于KMS激活Windows的官方说明文档](https://docs.microsoft.com/zh-cn/windows-server/get-started/kmsclientkeys)
通过`set`指令设置各版本激活秘钥以方便在下一步执行，KMS密钥请见微软官方说明文档
#### 3)激活
从注册表中读取Windows版本以选择合适的密钥，如果版本信息吻合则会进行激活操作，反之则判断为不支持VOL的版本
**注意：本操作不会卸载密钥和KMS服务器，而是直接进行覆盖，如有Bug请选择清除KMS激活后重启再试！**
### 2.Office的激活
#### 1)检测Office版本
通过`call`指令调用获取Office目录的分支并输出Office的目录给激活Office的分支
#### 2)获取Office目录
通过`set`参数定义32位系统和64位系统的Office目录，然后用`if`指令判断是32位还是64位并输出对应目录的路径
#### 3)设置激活密钥（Office2016-2019）
[Microsoft关于KMS激活Office的官方说明文档](https://docs.microsoft.com/zh-cn/deployoffice/vlactivation/gvlks)
通过识别Office目录中的文件内容来确定许可证版本，并根据版本输入合适密钥给予下一步
#### 4)激活Office
获取到Office目录并通过`cd`更改提示符的运行路径，运行Office目录下的`ospp.vbs`来使用KMS激活密钥激活Office
## 二、手动激活
### 1.激活Windows
#### 尚未完成，更新中……
#### Not finished yet, updating ...
