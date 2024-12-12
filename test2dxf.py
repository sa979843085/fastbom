import ezdxf

# 读取现有的 DXF 文件
doc = ezdxf.readfile('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/WHJ82蒸发波导诊断系统框图A3.dxf')


# 获取图纸空间布局
layout = doc.layout('Layout1')

print(layout)

# 查询视口实体
viewports = layout.query('VIEWPORT')

print(viewports)

for viewport in viewports:
    # 获取视口属性
    center_point = viewport.dxf.center
    height = viewport.dxf.height
    width = viewport.dxf.width

    # 打印视口属性
    print(f'Viewport center point: {center_point}')
    print(f'Viewport height: {height}')
    print(f'Viewport width: {width}')

# 在视口中添加一个矩形
# layout.add_lwpolyline([(0, 0), (10, 0), (10, 10), (0, 10), (0, 0)], dxfattribs={'layer': 'LWPOLYLINE'})

# 保存修改后的 DXF 文件
doc.saveas('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/WHJ82蒸发波导诊断系统框图A3_modified.dxf')


