# Python 办公自动化(Excel篇)

### 1. 数据盘点统计
#### 任务描述
> 通过数据盘点表中的数据，将数据定标表所需要的列信息进行统计后，生成对应表格文件

`任务列表:`
1. 从数据盘点表 **`test/test.xlsx`** 中筛选出 **`config.py中FILTERED_TABLE_HEADER`** 指定列中的数据
2. 对 **`区域`** 进行分组，之后再将各 **`区域`** 中的 **`字段名`** 进行分组，统计 **`各区域`** 中的 **`字段名`** ，但 **`只统计相同区域的相同字段名`**
3. 最后将分组统计后的数据导出为Excel文件