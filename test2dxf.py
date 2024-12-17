import ezdxf
from ezdxf import bbox
# 读取现有的 DXF 文件
doc = ezdxf.readfile('./WHKTA3.dxf')
msp = doc.modelspace()

# 获取模型空间中所有实体的边界框
entities = list(msp)
extents = bbox.extents(msp)

# 计算边界框的中心点
center_point = extents.center

print(f"模型空间的中心点坐标为: {center_point}")


# # 获取图纸空间布局
# layout = doc.layout('模型')

# # 查询视口实体
# viewports = layout.query('VIEWPORT')

# for viewport in viewports:
#     # 获取视口属性
#     center_point = viewport.dxf.center
#     height = viewport.dxf.height
#     width = viewport.dxf.width

#     # 打印视口属性
#     print(f'Viewport center point: {center_point}')
#     print(f'Viewport height: {height}')
#     print(f'Viewport width: {width}')

# # 在视口中添加一个矩形
# layout.add_lwpolyline([(0, 0), (10, 0), (10, 10), (0, 10), (0, 0)], dxfattribs={'layer': 'LWPOLYLINE'})

# # 保存修改后的 DXF 文件
# doc.saveas('./WHJ82蒸发波导诊断系统框图A3_modified.dxf')


