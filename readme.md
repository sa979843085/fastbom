## 介绍

我也不知道应该写啥

这个项目很简单，我却写了好几天

就是BOM表格的快速整理


~~然后分别导入word和cad中~~(并没有实现)

### 框图探索

#### 前提

假设我现在有一个A3的模板，在脚本的相同路径之中，名字叫WHKTA3.dxf

#### 用到的模块

1. ezdxf
2. pandas

#### 思路

1. 读取表格，获取表格内容并整理
    1. 获取表格
2. 读取框图，获取框图内容
    1. 获取框图的尺寸坐标
    2. 获取框图中的文字内容
3. 将表格内容新建表格填入框图模板中
    1. 内容要求
        1. 表格的内容需要按父件的代号分为多个小表格
    2. 格式要求
        1. 表格需要画出表格线
        2. 文字需要在表格的上下居中，左右靠左的位置，文字高度：2.5，字体：仿宋_GB2312.ttf
        3. 表格需要和框图对齐
    3. 位置要求
        1. 表格需要在图框内
        2. 各个小表格之间需要有一定的间隔
        3. 表格的行数需要根据内容自动调整
4. 保存框图

#### 实现细节

1. 读取表格为bom_data
2. 将表格分组为0级BOM、1级BOM、2级BOM第1组、2级BOM第二组...
    1. 单独处理0级BOM
        1. 当阶层为0时，这一行为0级BOM，将这一行单独存储然后从df中删除
    2. 处理1级及以上的BOM，将bom_data根据父件的代号后四位建立两列参考，第一列是bom层级，第二列是层级内的序号
        1. 当后四位均为0时为1级BOM，bom层级列是1，序号列是0
        2. 当后四位中有三个0时，为2级BOM，bom层级列是2，序号列是四位数从右往左第一个非0的数
        3. 当后四位中有两个0时，为3级BOM，bom层级列是3，序号列是四位数从右往左第一个非0的数
        4. 当后四位中有一个0时，为4级BOM，bom层级列是4，序号列是四位数从右往左第一个非0的数
        5. 当后四位全不为0时，为5级BOM，bom层级列是5，序号列是四位数从右往左第一个非0的数
    3. 将bom层级和序号进行合并，bom层级是整数部分，序号是小数部分，形式是1.0、2.1、3.2
    4. 将bom_data根据合并列进行分组
3. 读取框图，获取框图内容

