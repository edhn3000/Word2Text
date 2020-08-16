Word2Text
==================
## 介绍
  word文档转为txt，或反向转换


## 使用方式
  参照`Run.bat`，修改输入路径、输出路径等参数，并执行。
  或直接使用`java -jar Word2Text ...`方式执行
更多参数：

```sh
命令语法：
	java -jar Word2Text path [-encoding=] [-update] [-outdir=]
参数说明：
    path        源文件所在路径路径，注意结尾不要带\
    -encoding   转换后输出文件的字符集，如GBK、UTF-8等
    -update     已转换过且word文档未被更新的时不重复转换
    -outdir     输出路径，默认是跟原文件同目录
    -txt2docx   反转换，文本转word文档（目前仅支持生成docx）
```