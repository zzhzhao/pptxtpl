#!/usr/bin/env python
#-*- coding:utf-8 -*-
import json
import pptxtpl
from pptx.util import Cm

# example1:
# 现在有若干学生的信息，需要在ppt上渲染

pptx_obj = pptxtpl.PPTXTemplate("./example.pptx")


def get_replace_data(data):
    replace_data = {}
    for index, item in enumerate(data):
        for key, value in item.items():
            replace_data[pptx_obj.add_ppt_label(key + str(index))] = value

    return replace_data

def get_table_data(data):
    headers = list(data[0].keys())
    table_data = []
    for index, item in enumerate(data):
        line = [index + 1]
        for header in headers:
            line.append(item[header])

        table_data.append(line)

    return table_data


# 一班数据
data1 = [
    {"name": "zzz", "age": 90},
    {"name": "wb", "age": 45}]

replace_data1 = get_replace_data(data1)
# 学生人数
replace_data1["{student_number}"] = len(data1)

# print(json.dumps(replace_data1, indent=4, ensure_ascii=False))


# 替换ppt模板中 的六边形的数据
pptx_obj.replace_data(0, replace_data1)
# 删除其中未被渲染的数据
pptx_obj.delete_shapes(0)
# 获取两个组合图形
group_shape = pptx_obj.get_slide_group_shapes(0)

# Cm 指的是厘米，数据是幻灯片中组合图形摆放正确位置后获取的
size_list = [{"left": Cm(5.76)}, {"left": Cm(13.89)}]
# 将两个组合图形移动到中间
pptx_obj.update_group_shape_position_size(group_shape, size_list)


# 替换ppt模板中 的表格数据，如果有多个表格，会随机选择一个填充
# 遇到多个表格，可以将table[0][0]中做一个标签，以便唯一定位该表格
table_data = get_table_data(data1)
pptx_obj.add_table_data(index=0, data=table_data)


# 替换ppt模板中柱状图的数据
title_data = {"{grade_title}": {"category": ["不及格", "及格"], "data": {"一班": [20, 80], "二班": [30, 80]}}}
title_replace = {"{grade_title}": "一班二班及格人数柱状图"}
pptx_obj.replace_bar_chart_data(index=1, title_data=title_data, title_replace=title_replace)



# 若干班级的学生数据，幻灯片形式相同，如何渲染？
# 办法一：利用pptx_copy_slide复制幻灯片

# 由于幻灯片0 的数据已经渲染，复制时，就会复制渲染后的幻灯片
pptx_obj.pptx_copy_slide(0, 1)


# 办法二：复制方法中无法复制带有chart的统计图，比如柱状图，一般可在模板中手动复制若干个幻灯片
#        然后将没有填充的幻灯片利用删除方法删除多余幻灯片即可



pptx_obj.save("./save_example.pptx")



