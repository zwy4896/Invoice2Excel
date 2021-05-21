# Invoice2Excel

摘要：这篇文章介绍如何把发票内容提取出来保存到Excel中。

------

### 程序功能

程序会把发票中的内容提取出来然后写入Excel中，一个示例的发票内容如下：

![发票示例](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/202004/demo.PNG)

提取结果如下：

![提取结果](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/202004/result.png)

### 程序使用

- 下载代码

  ```shell
  git clone https://gitee.com/yczha/Invoice2Excel.git
  ```

- 安装依赖包：

  ```shell
  python -m pip install -r requirements.txt
  ```
  
- 准备数据

  准备好PDF文件，放置到一个目录下

- 运行程序

  ```shell
   # 注意：这里data指你的pdf文件放置的文件夹，参考example文件夹的结构 -p data也可以替换为--path=data
  python Invoice2Excel.py --path=data
  ```

### 更多

- 运行测试，可以通过以下命令运行测试

  ```shell
  python Invoice2Excel.py
  ```

- debug模式：会显示example/test.pdf文件的抽取情况，可视化展示

```shell
python Invoice2Excel.py --debug
```
结果如下：

![线段补全](http://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/202007/Figure_1.png)

![拆解单元格](http://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/202007/Figure_2.png)

![单词放入单元格](http://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/202007/Figure_3.png)


- 指定输出文件位置

  ```shell
  # 注意：这里data.xlsx指你的pdf结果保存文件， -s data.xlsx也可以替换为--save=data.xlsx
  python Invoice2Excel.py -s data.xlsx
  # 也可以同时指定两个参数
  python Invoice2Excel.py -s data.xlsx -path=data
  ```

### 获取帮助

联系作者获取帮助：

- 微信：18217235290
- Email：yooongchun@foxmail.com
- 博客：[永春小站](http://www.yooongchun.com)
