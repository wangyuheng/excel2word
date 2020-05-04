# excel2word

excel中每条记录生成一个word文档

## 操作

1. 准备word模板，变量设置为 `${xx}`
2. word另存为xml
    2.1 check 导出的xml文件变量格式是否正确(部分隐藏格式会导致变量中插入标签)
3. 准备excel，header对应变量名
4. 执行程序