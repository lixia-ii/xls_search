因为直接打包了python3.8的环境到svn目录里，所以pip的使用需要用如下形式
在common\tools\python-3.8.2\目录下，调用python -m pip install xxx

log:
1 为了适配打包python3.8的环境到svn目录里，手动删除了xlwings-script.py文件里原来带的python路径提示信息