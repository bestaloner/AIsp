#!/usr/bin/env python3
"""Generate a draw.io mind map from 数据集模板.xlsx — for Feishu import."""

import json
import xml.sax.saxutils as saxutils
import re

def xml_safe(text):
    """Sanitize text for XML attribute/value embedding."""
    # Remove control characters not allowed in XML 1.0
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    # Escape XML special chars
    return saxutils.escape(text, {'"': '&quot;'})

# Load the data
with open(r'f:\0.AI设计库\ai视频识别\gh-pages-deploy\template_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Track edge/cell IDs
cell_id = 2
def next_id():
    global cell_id
    cell_id += 1
    return str(cell_id)

# Build XML elements
cells = []
edges = []

# Styles
ROOT_STYLE = "rounded=1;whiteSpace=wrap;html=1;fillColor=#2E41CD;strokeColor=#1A2D99;fontColor=#FFFFFF;fontSize=16;fontStyle=1;arcSize=20;"
LEVEL1_STYLE = "rounded=1;whiteSpace=wrap;html=1;fillColor=#EFF6FF;strokeColor=#2563EB;fontSize=13;fontStyle=1;arcSize=12;fontColor=#1E40AF;"
LEVEL2_STYLE = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FEF3C7;strokeColor=#F57A00;fontSize=11;arcSize=10;fontColor=#92400E;align=left;spacingLeft=4;"
LEVEL3_STYLE = "rounded=0;whiteSpace=wrap;html=1;fillColor=#F0FDF4;strokeColor=#10B981;fontSize=10;arcSize=6;fontColor=#065F46;align=left;spacingLeft=3;"
LEVEL4_STYLE = "rounded=0;whiteSpace=wrap;html=1;fillColor=#FAFAFA;strokeColor=#D9D9D9;fontSize=9;arcSize=4;fontColor=#595959;align=left;spacingLeft=2;"

# Root node
ROOT_X, ROOT_Y = 40, 350
root_id = "2"
cells.append(f'<mxCell id="{root_id}" value="{xml_safe("数据集模板")}&#xa;{xml_safe("12类隐患·四级联动")}" style="{ROOT_STYLE}" vertex="1" parent="1"><mxGeometry x="{ROOT_X}" y="{ROOT_Y}" width="150" height="60" as="geometry"/></mxCell>')

# Layout constants
LEVEL_DX = 260
ITEM_DY = 90
DESC_DY = 32
CLAUSE_DY = 24
TEXT_DY = 20

cur_y = 60  # starting Y for level 1 items

for item_name, descs in data.items():
    item_id = next_id()
    item_x = ROOT_X + LEVEL_DX
    item_y = cur_y

    # Count total leaf nodes to size this node
    total_clauses = sum(len(clauses) for clauses in descs.values())
    item_h = max(50, len(descs) * 80 + 10)

    # Level 1: 检查项目
    short_name = item_name[:12] + ('…' if len(item_name) > 12 else '')
    cells.append(f'<mxCell id="{item_id}" value="{xml_safe(item_name)}" style="{LEVEL1_STYLE}" vertex="1" parent="1"><mxGeometry x="{item_x}" y="{item_y}" width="180" height="{item_h}" as="geometry"/></mxCell>')

    # Edge: root → item
    edge_id = next_id()
    edges.append(f'<mxCell id="{edge_id}" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#2563EB;strokeWidth=2;endArrow=classic;endFill=1;" parent="1" source="{root_id}" target="{item_id}" edge="1"><mxGeometry relative="1" as="geometry"/></mxCell>')

    desc_y = item_y + 5
    for desc_name, clauses in descs.items():
        desc_id = next_id()
        desc_x = item_x + LEVEL_DX

        # Truncate long descriptions
        desc_label = desc_name[:50] + ('…' if len(desc_name) > 50 else '')

        # Count total texts
        total_texts = sum(len(texts) for texts in clauses.values())
        desc_h = max(32, len(clauses) * 70 + 10)

        cells.append(f'<mxCell id="{desc_id}" value="{xml_safe(desc_label)}" style="{LEVEL2_STYLE}" vertex="1" parent="1"><mxGeometry x="{desc_x}" y="{desc_y}" width="280" height="{desc_h}" as="geometry"/></mxCell>')

        # Edge: item → desc
        edge_id = next_id()
        edges.append(f'<mxCell id="{edge_id}" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#F57A00;strokeWidth=1.5;endArrow=classic;endFill=1;" parent="1" source="{item_id}" target="{desc_id}" edge="1"><mxGeometry relative="1" as="geometry"/></mxCell>')

        clause_y = desc_y + 3
        for clause_name, texts in clauses.items():
            clause_id = next_id()
            clause_x = desc_x + LEVEL_DX
            clause_label = clause_name[:45] + ('…' if len(clause_name) > 45 else '')
            clause_h = max(24, len(texts) * 55 + 8)

            cells.append(f'<mxCell id="{clause_id}" value="{xml_safe(clause_label)}" style="{LEVEL3_STYLE}" vertex="1" parent="1"><mxGeometry x="{clause_x}" y="{clause_y}" width="280" height="{clause_h}" as="geometry"/></mxCell>')

            # Edge: desc → clause
            edge_id = next_id()
            edges.append(f'<mxCell id="{edge_id}" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#10B981;strokeWidth=1;endArrow=classic;endFill=1;" parent="1" source="{desc_id}" target="{clause_id}" edge="1"><mxGeometry relative="1" as="geometry"/></mxCell>')

            text_y = clause_y + 2
            for text_val in texts:
                text_id = next_id()
                text_x = clause_x + LEVEL_DX
                text_label = text_val[:60] + ('…' if len(text_val) > 60 else '')

                cells.append(f'<mxCell id="{text_id}" value="{xml_safe(text_label)}" style="{LEVEL4_STYLE}" vertex="1" parent="1"><mxGeometry x="{text_x}" y="{text_y}" width="320" height="20" as="geometry"/></mxCell>')

                # Edge: clause → text
                edge_id = next_id()
                edges.append(f'<mxCell id="{edge_id}" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#D9D9D9;strokeWidth=0.75;endArrow=classic;endFill=1;" parent="1" source="{clause_id}" target="{text_id}" edge="1"><mxGeometry relative="1" as="geometry"/></mxCell>')

                text_y += TEXT_DY + 2

            clause_y += clause_h + 8

        desc_y += desc_h + 8

    cur_y += item_h + 20

# Calculate total page size
total_w = 40 + LEVEL_DX * 4 + 340
total_h = max(cur_y + 40, 800)

xml = f'''<mxfile host="Electron" agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) draw.io/27.0.9 Chrome/134.0.6998.205 Electron/35.4.0 Safari/537.36" version="27.0.9">
  <diagram name="数据集模板-思维导图" id="mindmap-template">
    <mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="{total_w}" pageHeight="{total_h}" math="0" shadow="0">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
        <mxCell id="legend" value="&lt;b&gt;图例&lt;/b&gt;&lt;br&gt;🟦 检查项目 (6项)&lt;br&gt;🟧 事实描述&lt;br&gt;🟩 规定条款&lt;br&gt;⬜ 规定原文" style="text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=top;whiteSpace=wrap;fontSize=10;fontColor=#8C8C8C;" vertex="1" parent="1">
          <mxGeometry x="{total_w-200}" y="40" width="160" height="80" as="geometry" />
        </mxCell>
{chr(10).join(cells)}
{chr(10).join(edges)}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>'''

output = r'f:\0.AI设计库\ai视频识别\现版本设计文件\diagrams\05_数据集模板思维导图.drawio'
with open(output, 'w', encoding='utf-8') as f:
    f.write(xml)

print(f'Generated: {output}')
print(f'Total cells: {cell_id}')
print(f'Page size: {total_w}x{total_h}')
