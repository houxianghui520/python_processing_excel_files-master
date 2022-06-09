# python处理excel文件

#### 介绍
使用了pandas和xlwings处理excel文件.
为了更好的兼顾效率和兼容性, 使用pandas对数据进行处理,采用xlwings 进行填充 

#### 软件架构
软件架构说明


#### 安装教程

1.  xxxx
2.  xxxx
3.  xxxx

#### 使用说明

1.  pandas在批量删除时应先将待删除行加入列表,而后统一删除列表行,采用循环语句逐行删除会出现错误.
2.  pandas无法将结果写入已有excel的sheet中, 采用pandas+openpyxl可以将数据导出到已有的excel中的新sheet中.
3.  本例中将pandas处理后的数据导出到output.xlsx文件中,而后由xlwings将output中内容与目标文件合并,可以采用嵌套列表写入单元格的方法实现批        量写入,效率较高.

#### 参与贡献

1.  Fork 本仓库
2.  新建 Feat_xxx 分支
3.  提交代码
4.  新建 Pull Request


#### 码云特技

1.  使用 Readme\_XXX.md 来支持不同的语言，例如 Readme\_en.md, Readme\_zh.md
2.  码云官方博客 [blog.gitee.com](https://blog.gitee.com)
3.  你可以 [https://gitee.com/explore](https://gitee.com/explore) 这个地址来了解码云上的优秀开源项目
4.  [GVP](https://gitee.com/gvp) 全称是码云最有价值开源项目，是码云综合评定出的优秀开源项目
5.  码云官方提供的使用手册 [https://gitee.com/help](https://gitee.com/help)
6.  码云封面人物是一档用来展示码云会员风采的栏目 [https://gitee.com/gitee-stars/](https://gitee.com/gitee-stars/)
