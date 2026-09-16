# 1. 用 part_related_by 查询关系，而不是属性访问
try:
    part = doc.part.part_related_by(RT_NUMBERING)
except KeyError:
    part = None

# 2. 找不到才新建，用 XmlPart 而不是 NumberingPart
from docx.opc.part import XmlPart
new_part = XmlPart(partname, content_type, element, doc.part.package)
doc.part.relate_to(new_part, RT_NUMBERING)
