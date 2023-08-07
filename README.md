<p align="center">
	👉 <a target="_blank" href="https://www.python-office.com/">项目官网：https://www.python-office.com/</a> 👈
</p>
<p align="center">
	👉 <a target="_blank" href="https://python-office-1300615378.cos.ap-chongqing.myqcloud.com/python-office.jpg">本开源项目的交流群</a> 👈
</p>



-------------------------------------------------------------------------------

## 📚简介

wftools是python自动化办公的小工具的代码合集。

-------------------------------------------------------------------------------

## 📦安装

### 🍊pip 自动下载&更新

```
pip install -i https://mirrors.aliyun.com/pypi/simple/ poexcel -U
```

-------------------------------------------------------------------------------

## 📝功能

[📘官网：https://www.python-office.com/](https://www.python-office.com/)

| 序号 | 方法名             | 功能                              | 视频（文档）                                                                | 演示代码                                                                                                                                                                                                                                            |
|----|-----------------|---------------------------------|-----------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 1  | fake2excel      | 批量生成Excel数据                     | [视频](https://www.bilibili.com/video/BV1wr4y1b7uk)                     | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/%E6%89%B9%E9%87%8F%E6%A8%A1%E6%8B%9F%E6%95%B0%E6%8D%AE.py)                                                                                                          |
| 2  | merge2excel     | 合并多个Excel到一个Excel的不同sheet中      | [视频](https://www.bilibili.com/video/BV1714y147Ao)                     | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/%E5%90%88%E5%B9%B6%E5%A4%9A%E4%B8%AAExcel%E5%88%B0%E4%B8%80%E4%B8%AAExcel%E7%9A%84%E4%B8%8D%E5%90%8Csheet%E4%B8%AD.py)                                              |
| 3  | sheet2excel     | 同一个excel里的不同sheet，拆分为不同的excel文件 | [视频](https://www.bilibili.com/video/BV1714y147Ao)                     | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/%E5%90%8C%E4%B8%80%E4%B8%AAexcel%E9%87%8C%E7%9A%84%E4%B8%8D%E5%90%8Csheet%EF%BC%8C%E6%8B%86%E5%88%86%E4%B8%BA%E4%B8%8D%E5%90%8C%E7%9A%84excel%E6%96%87%E4%BB%B6.py) |
| 4  | find_excel_data | 根据内容查询Excel                     | [视频](https://www.bilibili.com/video/BV1Bd4y1B7yr)                     | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/%E6%A0%B9%E6%8D%AE%E5%86%85%E5%AE%B9%EF%BC%8C%E6%9F%A5%E8%AF%A2Excel.py)                                                                                            |
| 5  | excel2pdf       | Excel转PDF                       | [视频](https://www.bilibili.com/video/BV1A84y1N7or)                     | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/Excel%E8%BD%ACPDF.py)                                                                                                                                               |
| 6  | query4excel     | 把100个Excel中符合条件的数据，汇总到1个Excel里  | [视频](https://www.bilibili.com/video/BV1Hs4y1S7TT)                     | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/%E6%8A%8A100%E4%B8%AAExcel%E4%B8%AD%E7%AC%A6%E5%90%88%E6%9D%A1%E4%BB%B6%E7%9A%84%E6%95%B0%E6%8D%AE%EF%BC%8C%E6%B1%87%E6%80%BB%E5%88%B01%E4%B8%AAExcel%E9%87%8C.py)  |
| 7  | count4page      | 统计Excel打印出来有多少页                 | [文档](https://blog.csdn.net/weixin_42321517/article/details/131218163) | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/%E7%BB%9F%E8%AE%A1Excel%E6%89%93%E5%8D%B0%E5%87%BA%E6%9D%A5%E6%9C%89%E5%A4%9A%E5%B0%91%E9%A1%B5.py)                                                                 |
| 8  | merge2sheet      | 统计Excel打印出来有多少页                 | [文档](https://blog.csdn.net/weixin_42321517/article/details/131218163) | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/合并2个Excel的内容到一个sheet中.py)                                                                 |
| 9  | split_excel_by_column      | 根据指定的列，拆分Excel到不同的sheet         | [文档](https://blog.csdn.net/weixin_42321517/article/details/131218163) | [源码](https://github.com/CoderWanFeng/python-office/blob/master/demo/poexcel/根据指定的列，拆分excel.py)                                                                 |

## 🏗️添砖加瓦

### 📐PR的建议

python-office欢迎任何人来添砖加瓦，贡献代码，建议提交的pr（pull request）符合一些规范，规范如下：

参与项目建设的步骤：

- 例如：你需要给python-office添加一个add方法。
    1. 你的Github账户名为：demo
    2. 于是你在./contributors新建了文件夹./demo
    3. 新建了add.py文件，编辑你的代码
    4. 编辑完成，提交pr到master分支（gitee或者GitHub，都可以）。可以注明你对自己功能的取名建议
    5. 晚枫收到后，会对各位的代码进行测试后，合并后打包上传到python官方库

### 📐代码规范

1. 注释完备，尤其每个新增的方法应按照Google Python文档规范标明方法说明、参数说明、返回值说明等信息，必要时请添加单元测试，如果愿意，也可以加上你的大名。
2. python-office的文档，需要进行格式化。注意：只能格式化你自己的代码
3. 请直接pull request到`master`分支。`master`是主分支，表示已经发布pypi库的版本。**未来参与人数增多，会开辟新的分支，请留意本文档的更新。
   **
4. 我们如果关闭了你的issue或pr，请不要诧异，这是我们保持问题处理整洁的一种方式，你依旧可以继续讨论，当有讨论结果时我们会重新打开。

### 🧬贡献代码的步骤

1. 在Gitee或者Github上fork项目到自己的repo
2. 把fork过去的项目也就是你的项目clone到你的本地
3. 修改代码
4. commit后push到自己的库
5. 登录Gitee或Github在你首页可以看到一个 pull request 按钮，点击它，填写一些说明信息，然后提交到master分支即可。
6. 等待维护者合并

### 🐞提供bug反馈或建议

提交问题反馈时，请务必填写和python-office代码本身有关的问题，不进行有关python学习，甚至是个人练习的知识答疑和讨论。

- [Github issue](https://github.com/CoderWanFeng/poexcel/issues)

-------------------------------------------------------------------------------

## 📌联系作者

<p align="center" id='开源交流群-banner'>
<a target="_blank" href='https://python-office-1300615378.cos.ap-chongqing.myqcloud.com/python-office.jpg'>
<img src="https://python-office-1300615378.cos.ap-chongqing.myqcloud.com/python-office-qr.jpg" width="100%"/>
</a> 
</p>