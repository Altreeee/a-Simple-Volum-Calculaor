'''
接收一个exel表格，根据给的两个数据量确定新的数据断点量，
使所有数据可以两两对应
直接使用比例关系确定没有数据的断点
使用新断点得到两曲线所有两两对应的点的高度差，和两断点间距离
求两断点间梯形体积
相加
'''
import pandas as pd
import numpy as np
import math
#引入exel表格，其中左数第三列第2行开始为下线横坐标，左数第四列第2行开始为下线纵坐标
#左数第六列第2行开始为上线横坐标，左数第七列第2行开始为上线纵坐标
#左数第9列上数第2行为边坡

#首先使用两组数据保存两条线的坐标。例：[上横坐标，上纵坐标]，[下横坐标，下纵坐标]
#然后以横坐标为基准将两组坐标合并为[横坐标，上纵坐标，下纵坐标],无法确定的纵坐标数字使用未知数x代替，
#保证横坐标能够覆盖上下两线的所有横坐标
#得到一组新数据，可以看到所有表格中有的横坐标点的数据


# 读取Excel文件
excel_file = '数据填写表格.xlsx'
df = pd.read_excel(excel_file)

upper_base = df.iloc[0, 12]
# 读取起点横坐标、终点横坐标和对应的bianpo值
bianpo_data = df.iloc[:, [8, 9, 10]].values

# 处理bianpo值
def set_bianpo(x_coord):
    for row in bianpo_data:
        start_coord, end_coord, bianpo = row
        start_coord = float(start_coord)  # Convert to float
        end_coord = float(end_coord)      # Convert to float
        #print(start_coord)
        if start_coord <= x_coord <= end_coord:
            return bianpo
    return None


# 提取下线和上线的坐标数据
lower_coords = df.iloc[:, [2, 3]].values
upper_coords = df.iloc[:, [5, 6]].values

# 合并坐标数据
all_x_coords = sorted(set(lower_coords[:, 0]) | set(upper_coords[:, 0]))

merged_coords = []
for x_coord in all_x_coords:
    lower_coord = lower_coords[lower_coords[:, 0] == x_coord, 1]
    upper_coord = upper_coords[upper_coords[:, 0] == x_coord, 1]
    
    if len(lower_coord) > 0:
        y_lower = lower_coord[0]
    else:
        y_lower = np.nan
        
    if len(upper_coord) > 0:
        y_upper = upper_coord[0]
    else:
        y_upper = np.nan
    
    # 获取对应的bianpo值
    x_bianpo = set_bianpo(x_coord)
        
    merged_coords.append([x_coord, y_upper, y_lower, x_bianpo])

# 删除所有纵坐标都为未知的数据
filtered_coords = [coord for coord in merged_coords if not (np.isnan(coord[1]) and np.isnan(coord[2]))]

# 按照横坐标从小到大排序
sorted_coords = sorted(filtered_coords, key=lambda x: x[0])


# 求解未知的 nan 值
for i in range(len(sorted_coords)):
    if np.isnan(sorted_coords[i][1]):
        prev_known = None
        next_known = None
        
        # 找到前一个已知数据点
        for j in range(i - 1, -1, -1):
            if not np.isnan(sorted_coords[j][1]):
                prev_known = sorted_coords[j]
                break
        
        # 找到后一个已知数据点
        for j in range(i + 1, len(sorted_coords)):
            if not np.isnan(sorted_coords[j][1]):
                next_known = sorted_coords[j]
                break
        
        # 如果找到了前后两个已知数据点
        if prev_known is not None and next_known is not None:
            x_prev, y_prev = prev_known[0], prev_known[1]
            x_next, y_next = next_known[0], next_known[1]
            x_curr = sorted_coords[i][0]
            
            y_estimated = ((x_curr - x_prev) / (x_next - x_prev)) * (y_next - y_prev) + y_prev
            sorted_coords[i][1] = y_estimated

# 求解未知的第三列数据
for i in range(len(sorted_coords)):
    if np.isnan(sorted_coords[i][2]):
        prev_known = None
        next_known = None
        
        # 找到前一个已知数据点
        for j in range(i - 1, -1, -1):
            if not np.isnan(sorted_coords[j][2]):
                prev_known = sorted_coords[j]
                break
        
        # 找到后一个已知数据点
        for j in range(i + 1, len(sorted_coords)):
            if not np.isnan(sorted_coords[j][2]):
                next_known = sorted_coords[j]
                break
        
        # 如果找到了前后两个已知数据点
        if prev_known is not None and next_known is not None:
            x_prev, y_prev = prev_known[0], prev_known[2]
            x_next, y_next = next_known[0], next_known[2]
            x_curr = sorted_coords[i][0]
            
            y_estimated = ((x_curr - x_prev) / (x_next - x_prev)) * (y_next - y_prev) + y_prev
            sorted_coords[i][2] = y_estimated


# 计算体积
volumes = []
filled_coords = sorted_coords
print(len(filled_coords))
for i in range(len(filled_coords) - 1):
    x1, y1, z1, bianpo1 = filled_coords[i]
    x2, y2, z2, bianpo2 = filled_coords[i + 1]
    
    # 计算截面高度、上底、下底
    h1 = z1 - y1
    lower_base1 = upper_base + h1 * bianpo1 * 2
    
    h2 = z2 - y2
    lower_base2 = upper_base + h2 * bianpo2 * 2
 
    
    # 计算截面面积和体积
    area1 = h1 * (upper_base + lower_base1) / 2
    area2 = h2 * (upper_base + lower_base2) / 2
    distance = x2 - x1
    volume = (area1 + area2 + math.sqrt(area1 * area2)) * distance / 3
    volumes.append(volume)

# 计算总体积
total_volume = sum(volumes)

# 打印各个截面的体积
for i, volume in enumerate(volumes):
    print(f"Volume between {filled_coords[i]} and {filled_coords[i+1]}: {volume}")

# 打印总体积
print("Total Volume:", total_volume)

input("按下回车键以退出...")