# 数据录入小工具（for 尹婧）

## 本工具有三份小代码：
- blankprep.py： 用于准备数据录入的空白excel表格
- sheetWriter.py: 用于将新的数据写入模板表格
- sheetWriter2.py: 手动录入信息并写入模板表格

## 使用工具之前：
- 将代码复制到目标文件夹
- 将“自动生成预算导入模板.xlsx” 也放入文件夹
- 用shift+右键，打开terminal，并打入python,出现python X.X.X 即表示运行正常

## 使用 blankprep.py
- 输入 python blankprep.py 并回车，
- 出现 “大概多少个款项录入?”字样，输入某个整数数字并回车即可（尽可能大），如果不输入，代码会自动默认为10
- 然后出现 “你想用的文件名?:”, 这里随便输入一些汉字，字母，均可，最终空白表格的名字会是“输入表-月日-你输入的字符.xlsx”
- 如果上一项不输入，表格名字默认为“输入表-月日-yj.xlsx”
- 回车之后，一个空白表格生成在同一文件夹。该表格总共有四列，出第一列有下拉菜单之外，其余列都可自由输入

## 手写输入信息之后，保存输入表， 开始使用 sheetWriter.py
- 输入 python sheetWriter.py 并回车，出现“你想用的文件名?:”
- 这里随便输入一些汉字，字母，均可，最终空白表格的名字会是"预算明细账模板-"+月日+输入的字符+".xlsx"
- 如果不输入，直接按回车，文件名将是"预算明细账模板-"+月日+“yj”+".xlsx"
- 回车后会出现 “备注要用到的信息? 比如，‘日常费用’”， 这里主要针对备注里面的东西，
- 如果直接回车，每一项的备注里会出现 “预拨+x月+机构名称+*日常费用*+月日-yj”
- 如果文件夹中有多个以“输出表” 开头的文件，代码会接着提示：“文件夹中存在多个输入文件，你需要哪一个?” 并显示文件名
- 如果第 n 个文件是你想用的，打一个 n （从 1 开始数）并回车即完成全部操作。

## 关于 sheetWriter2.py
- 类似于sheetwriter.py,但在录入信息时有以下提示

>请输入新的款项，格式为：类别 机构 科目 金额（注意空格）
>类别有3个选项，分别输入对应数字即可（1-固定-初始版,2-固定-调整版,3-变动-调整版）

- 在开始录入信息时，需要手动打入机构，科目，和金融的信息
- 如果没有新的信息录入，直接打回车即可，系统会提醒你总共录入了几个款项。

## 使用 Anaconda 运行代码
- 首先打开 Anaconda Navigator
- 找到 Anaconda powershell prompt 并点击，一个指令窗口将被打开
- 现在，找到代码所在的地址，这里我们用一个简单的地址作为例子：C:\user\12345
- 你也可以直接打开代码所在的文件夹，点击地址栏，直接复制地址(仅仅是地址，不包括文件名)
- 复制地址之后，我们利用上面的例子，先让prompt进入目标文件夹。 这时在指令窗口输入以下指令并回车：

>`cd C:\user\12345`

- 注意在`cd` 后面的空格
- 之后你可以使用以下指令使用代码：

> `python blankprep.py`

> `python sheetWriter.py`

## 注意事项
- 如果运行代码时出现错误，请在terminal 窗口中进行如下操作
- 输入 pip install openpyxl --upgrade 并回车
- 输入 pip install glob --upgrade 并回车
- 如果上述问题无法解决，请直接联系作者

**作者：Sizhe/03/11/20**