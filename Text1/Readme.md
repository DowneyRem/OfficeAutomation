## TEXT1
### 具体需求：
- 批量替换源文件内的坐标
- 输入 x0 后，批量替换所有 x 大于x0 的坐标 
- 在异常的 Python 环境下正常执行脚本
- 原始数据见 ResearchGraph.json


### 实现思路：
#### 编写python 脚本
1. 考虑到 ResearchGraph.json 不完全符合 json 格式的问题，不可使用 json 模块
2. 直接按照文本读取，使用正则直接替换，速度较上者应该更快

#### 执行 python 脚本
1. 将 python 脚本打包成 exe 程序，但可能受系统版本影响，无法正常运行
2. 使用便携版的 python 执行脚本

### 实现效果：
#### Replace.py
1. 读取文件后 ResearchGraph.json
2. 输入 x0 后，批量替换所有 x 大于x0 的坐标

![视图-分页预览.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Text1/运行.png)

#### Replace.cmd
1. 搜寻 Python 程序的路径
2. 拼合 Python 脚本的路径
3. 指定 Python 版本，执行脚本

![视图-分页预览.png](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Text1/程序等.png)
