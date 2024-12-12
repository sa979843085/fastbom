import ezdxf
from ezdxf.enums import TextEntityAlignment

# 创建一个新的 DXF 文档，指定 DXF 版本（如 R2013）
doc = ezdxf.new(dxfversion='R2013')

# 获取模型空间
msp = doc.modelspace()

# 创建文本样式
doc.styles.new('FangSong_GB2312', dxfattribs={'font': '仿宋_GB2312.ttf'})

# 添加文本并指定样式
text = msp.add_text("这是仿宋_GB2312字体的文本", height=1.5, dxfattribs={'style': 'FangSong'})

# 设置文本位置和对齐方式
text.set_placement((2, 6), align=TextEntityAlignment.LEFT)

doc.saveas('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/text_with_fangsong_font.dxf')