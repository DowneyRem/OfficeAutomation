## TEXT1
### 具体需求：
- 消除TXT乱码
- 将不同编码的TXT文件统一转成UTF8编码

### 实现思路：
1. 使用 chardet 模块，判断TXT编码后解码
2. 根据置信度高低分别处理
3. 高置信度的，重新编码为UTF8，再另存为TXT文件