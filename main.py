from fastapi import FastAPI, Response
from pydantic import BaseModel
from typing import List, Optional
import pandas as pd
import networkx as nx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from collections import defaultdict
import io

app = FastAPI()

# 定義從 Power Automate 接收的 JSON 資料格式
class CircuitData(BaseModel):
    PRODUCT_TYPE: str
    CIRCUIT_ID: str
    SITE_ID_A: str
    BUILDING_A: str
    SITE_ID_B: Optional[str] = ""
    BUILDING_B: Optional[str] = ""
    ROUTER_NAME: Optional[str] = ""

def apply_l2r_topology_layout(G):
    left_nodes = []
    center_nodes = []
    right_nodes = []
    for n, d in G.nodes(data=True):
        role = d.get('role', 'customer')
        if role == 'cloud': center_nodes.append(n)
        elif role == 'datacenter': right_nodes.append(n)
        else: left_nodes.append(n)

    pos = {}
    X_LEFT, X_CENTER, X_RIGHT = 2.0, 6.66, 11.33
    
    def assign_y_positions(nodes, x_pos):
        if not nodes: return
        y_spacing = 1.5
        start_y = 3.75 - ((len(nodes) - 1) * y_spacing / 2)
        for i, node in enumerate(nodes):
            pos[node] = [x_pos, start_y + (i * y_spacing)]

    assign_y_positions(left_nodes, X_LEFT)
    assign_y_positions(center_nodes, X_CENTER)
    assign_y_positions(right_nodes, X_RIGHT)
    return pos

@app.post("/generate-pptx")
def generate_pptx(data: List[CircuitData]):
    # 1. 將收到的 JSON 轉成 DataFrame
    df = pd.DataFrame([item.dict() for item in data]).fillna('')

    # 2. 建立 PPT
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shapes = slide.shapes

    G = nx.Graph()
    edges_info = defaultdict(list)

    # (套用您原本的繪圖與邏輯，此處略縮篇幅，請貼上您原本 for row in df.iterrows(): 以下的繪圖邏輯)
    for _, row in df.iterrows():
        site_a, site_b = str(row['SITE_ID_A']), str(row['SITE_ID_B'])
        p_type, c_id = str(row['PRODUCT_TYPE']).upper(), str(row['CIRCUIT_ID'])
        role_a = 'datacenter' if any(k in str(row['BUILDING_A']) for k in ['HQ', 'DC', '總部', '機房']) else 'customer'
        G.add_node(site_a, label=row['BUILDING_A'], role=role_a)
        circuit_label = f"{p_type} [{c_id}]"
        
        is_cloud_product = any(x in p_type for x in ['MPLS', 'VPN', 'ADSL'])
        if is_cloud_product:
            cloud_id = "Cloud_Network"
            G.add_node(cloud_id, label="MPLS / VPN / Internet", role='cloud')
            edges_info[tuple(sorted((site_a, cloud_id)))].append(circuit_label)
            if site_b:
                role_b = 'datacenter' if any(k in str(row['BUILDING_B']) for k in ['HQ', 'DC', '總部', '機房']) else 'customer'
                G.add_node(site_b, label=row['BUILDING_B'], role=role_b)
                edges_info[tuple(sorted((site_b, cloud_id)))].append(circuit_label)
        elif site_b: 
            role_b = 'datacenter' if any(k in str(row['BUILDING_B']) for k in ['HQ', 'DC', '總部', '機房']) else 'customer'
            G.add_node(site_b, label=row['BUILDING_B'], role=role_b)
            edges_info[tuple(sorted((site_a, site_b)))].append(circuit_label)

    layout_pos = apply_l2r_topology_layout(G)
    
    # ...(略，貼上您原本繪製連線與實體節點的程式碼)...

    # 3. 將生成的 PPT 存入記憶體中 (不寫入硬碟)
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)

    # 4. 回傳檔案給 Power Automate
    return Response(
        content=ppt_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=network_diagram.pptx"}
    )