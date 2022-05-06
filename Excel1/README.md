## Excel1

### 功能需求分析
#### Python 处理部分
1. (定期)从pixiv获取小说数据（小说名称、发布时间、**阅读量、收藏量**）
![获取数据](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/1.png)
2. 将数据格式化并保存至文件
![保存文件](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/2.png)


#### Excel VBA 处理部分
3. 将数据导入excel表格中
![导入数据](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/3.png)
4. 使用Excel进行简单的数据分析
![数据分析](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/4.png)


### 实现过程
1. 使用 pypixiv 获取小说数据
2. 使用 pandas 格式化数据，并导出xlsx
3. 使用 pywin32 打开Excel程序
4. 通过 Excel 函数引用，导入数据
5. 运行 VBA 代码，将获取的数据存放在数据源表中（【点】【赞】两个sheet）
6. 运行 VBA 代码，重新筛选【近期】sheet的数据
7. 运行 VBA 代码，更新【个人】sheet的数据
8. 使用 pywin32 将xlsm另存为xlsx
9. ~~使用 WPS Office，分享 xlsx文件~~



## 结果与效果
### 运行结果
0. 程序可以自动获取、格式化、处理数据，最终以excel表格的形式呈现出来
1. 节省每次手动录入数据以及筛选数据的时间，将原来的约15-20分钟，缩短至不到1分钟

### 最终效果

####  pixiv 数据分析
![pixiv 数据分析](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/0.png)



####  Excel 数据分析

#### 全部小说， 阅读量/收藏量的近期增量

![数据1](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/4.png)

#### 单篇小说 阅读量/收藏量的每周增量，及数据图

![数据2](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/5.png)

#### 全部小说  阅读量/收藏量分布

![数据3](https://raw.githubusercontent.com/DowneyRem/OfficeAutomation/main/Excel1/6.png)
