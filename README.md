# 活动发票数据整理工具（ActivityToolkit）

活动发票任务指的是：在奇绩的线上直播、线下创业交流会等活动前，给报名的用户发送门票。门票即为短信和邮件。发票所用平台为阿里云平台，此工具为发票前的数据整理工具。



## 主要功能

#### 1. 整合下述三个渠道的信息来源

- 麦克表单中的相应表单
- 微信群里和pr、活动组交接的报名名单
- 历史发送数据

#### 2. 解决两个数据整理需求（汇总 OR 去重）

- 发送数据整理=历史数据+新增数据
- 发送数据整理=新增数据=全部数据-历史数据

#### 3. 输出存档文件

由于发票任务时间跨度长，同一场活动的发票一般至少要发3次，时间跨度为一周。因此大概率同一活动的发票任务需要多人共同承担。这样的话就需要保存下每次的发票数据，方便之后负责发票的人参照。



## 使用方式

#### 【Step 1】安装R语言环境

https://mirrors.tuna.tsinghua.edu.cn/CRAN/

只需安装基础的R即可，不需要下载额外的IDE（例如Rstudio）

#### 【Step 2】第一次运行时请安装下列包

在console内依次输入

```R
install.packages('readxl')
install.packages('openxlsx')
install.packages('stringr')
```

#### 【Step 3】确定此次任务类型

如果是发新增数据，则下载使用Send New文件夹：

https://github.com/MiraclePlusBrain/ActivityToolkit/tree/main/Send%20All

如果是发历史全部数据，则下载使用Send All文件夹：

https://github.com/MiraclePlusBrain/ActivityToolkit/tree/main/Send%20New

#### 【Step 4】下载对应文件夹

并在文件夹内创建三个新的空文件夹（add，history，mike），命名如下：

![image-20220105115541497](https://github.com/MiraclePlusBrain/ActivityToolkit/blob/main/image-storage/1.png)

其中send_accumulate为核心代码，三个空文件夹用来存放需要合并和去重的数据

#### 【Step 5】获取今日发票数据至三个文件夹

##### 1、add文件夹

【add】文件夹存放从pr或者活动组那里获得的数据，示例如下：

![image-20220105143738341](https://github.com/MiraclePlusBrain/ActivityToolkit/blob/main/image-storage/2.png)

放入其中的excel表格需要满足几个条件：第一行需要是列名、列名需要包含“手机”和“邮箱“

##### 2、history文件夹

【history】文件夹存放由历史发票数据云盘下载的数据，示例如下：

<img src="https://github.com/MiraclePlusBrain/ActivityToolkit/blob/main/image-storage/3.png" alt="image-20220105145048929" style="zoom: 25%;" />

##### 3、mike文件夹

【mike】文件夹存放由麦克表单下载的数据，示例如下：

![image-20220105145205686](https://github.com/MiraclePlusBrain/ActivityToolkit/blob/main/image-storage/4.png)

放入其中的excel表格需要满足几个条件：第二行需要是列名（从麦克表单下载时这点**自然满足**）

#### 【Step 6】添加额外信息

打开add_info文件，向其中键入两个信息：

<img src="https://github.com/MiraclePlusBrain/ActivityToolkit/blob/main/image-storage/5.png" alt="image-20220105151030056" style="zoom:50%;" />

第一行和第二行都表示此次活动的地点。第一行的”南京“用于在麦克表单的数据中搜索相应城市的报名数据（从麦克表单下载的部分数据也包含其它活动场次的报名，所以需要筛选提取一下）。第二行用于给发票数据存档文件命名。

**若需要同时搜索两个城市的报名数据，可以将第一行改为”南京|广州“，即用|符号连接。*

#### 【Step 6】运行代码

双击打开send_accumulate.r或者send_new.r文件。**然后直接全部运行**。

若不知道怎么运行代码，也可以在windows系统中**点击”自动运行测试“文件**，自动执行这段代码。

#### 【Step 7】查看结果

运行代码之后，文件夹中会多出四个文件：

<img src="https://github.com/MiraclePlusBrain/ActivityToolkit/blob/main/image-storage/6.png" alt="image-20220105145929378" style="zoom:80%;" />

其中sendmail和sendtel可用于直接在阿里云平台上传（无需其他处理）

其中MPB_event_mail_XXXX_guangzhou_721和MPB_event_message_XXXX_guangzhou_720可用于**上传云盘存档**，便于下一个发票的人查看和使用。这两个文件的名字中中间的数字表示日期，最后的数字表示数据行数



## 注意事项

1. add文件夹中的所有文件第一行只能是列名
2. history文件夹中所有文件保证有两个sheet（如果历史文件都是根据这段代码生成的话，那么这里自然地没有问题）
3. 保证history文件夹中第一行没有列名
5. **所有表单需要都是xlsx文档**
