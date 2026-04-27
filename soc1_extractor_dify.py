#!/usr/bin/env python3
"""
SOC1 Report Content Extractor - Dify API Version
通过Dify平台调用LLM，提取SOC1报告内容
支持扫描件PDF（使用本地OCR pytesseract识别）
"""

import pdfplumber
import pandas as pd
import os
import json
import base64
import uuid
import sys
from pathlib import Path
from typing import Dict, List, Optional
import requests
from io import BytesIO
import pytesseract
from PIL import Image


def setup_tesseract_path():
    """
    配置Tesseract OCR路径
    打包后的exe需要指定Tesseract可执行文件位置
    """
    if sys.platform == 'win32':
        # 检查是否在打包环境中运行
        if getattr(sys, 'frozen', False):
            # 打包后的exe运行环境
            base_path = sys._MEIPASS
            tesseract_exe = os.path.join(base_path, 'tesseract', 'tesseract.exe')
            tessdata_path = os.path.join(base_path, 'tessdata')

            if os.path.exists(tesseract_exe):
                pytesseract.pytesseract.tesseract_cmd = tesseract_exe
                print(f"  [配置] 使用打包的Tesseract: {tesseract_exe}")

            if os.path.exists(tessdata_path):
                # 设置tessdata路径环境变量
                os.environ['TESSDATA_PREFIX'] = tessdata_path
                print(f"  [配置] 使用打包的tessdata: {tessdata_path}")
        else:
            # 开发环境：尝试自动检测系统安装的Tesseract
            tesseract_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            ]
            for path in tesseract_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    print(f"  [配置] 使用系统Tesseract: {path}")
                    break


# 初始化Tesseract路径
setup_tesseract_path()

# Dify API配置
API_KEY = "app-bHdz3erKJIqrqjd2funJ0eDS"
BASE_URL = "https://ai-platform-uat.ey.net/v1"

# 扫描件检测阈值（前5页文本总量低于此值判定为扫描件）
SCAN_THRESHOLD = 500

# 用户标识（用于Dify对话追踪）
USER_ID = "soc-extractor-" + str(uuid.uuid4())[:8]


def is_scanned_pdf(pdf_path: str) -> bool:
    """
    检测PDF是否是扫描件
    通过检查前5页的可提取文本量来判断
    """
    pdf = pdfplumber.open(pdf_path)
    total_text_length = 0
    total_pages = len(pdf.pages)

    for i in range(min(5, total_pages)):
        text = pdf.pages[i].extract_text() or ""
        total_text_length += len(text.strip())

    pdf.close()

    is_scanned = total_text_length < SCAN_THRESHOLD
    if is_scanned:
        print(f"  [检测] 该PDF为扫描件（前5页文本总量: {total_text_length} 字符）")
    else:
        print(f"  [检测] 该PDF为文本型（前5页文本总量: {total_text_length} 字符）")

    return is_scanned


def extract_pdf_text(pdf_path: str) -> str:
    """提取PDF全部文本"""
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)

    all_text = []
    for i in range(total_pages):
        page_text = pdf.pages[i].extract_text() or ""
        if page_text:
            all_text.append(f"[第{i+1}页]\n{page_text}")

    pdf.close()
    return "\n\n".join(all_text)


def extract_pdf_page_image(pdf_path: str, page_num: int, resolution: int = 150) -> str:
    """
    将PDF指定页面转换为base64编码的图像
    返回base64字符串
    """
    pdf = pdfplumber.open(pdf_path)
    if page_num < len(pdf.pages):
        page = pdf.pages[page_num]
        im = page.to_image(resolution=resolution)
        pil_image = im.original

        buffer = BytesIO()
        pil_image.save(buffer, format="PNG")
        img_bytes = buffer.getvalue()
        pdf.close()
        return base64.b64encode(img_bytes).decode('utf-8')
    pdf.close()
    return ""


def extract_all_pages_as_images(pdf_path: str, resolution: int = 150) -> List[str]:
    """
    将PDF所有页面转换为base64图像列表
    """
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)
    images = []

    for i in range(total_pages):
        page = pdf.pages[i]
        im = page.to_image(resolution=resolution)
        pil_image = im.original

        buffer = BytesIO()
        pil_image.save(buffer, format="PNG")
        img_bytes = buffer.getvalue()
        images.append(base64.b64encode(img_bytes).decode('utf-8'))

    pdf.close()
    return images


def ocr_pdf_pages(pdf_path: str, resolution: int = 200, lang: str = 'eng') -> str:
    """
    使用pytesseract对扫描件PDF进行OCR识别
    将每页转换为图像后进行OCR，合并为完整文本

    Args:
        pdf_path: PDF文件路径
        resolution: 图像分辨率（越高OCR效果越好，但速度更慢）
        lang: OCR语言（eng=英文，chi_sim=简体中文，chi_tra=繁体中文）

    Returns:
        OCR识别的完整文本，按页分段
    """
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)
    all_text = []

    print(f"  开始OCR识别（共{total_pages}页，分辨率{resolution}dpi，语言{lang}）...")

    for i in range(total_pages):
        page = pdf.pages[i]
        im = page.to_image(resolution=resolution)
        pil_image = im.original

        # OCR识别当前页面
        page_text = pytesseract.image_to_string(pil_image, lang=lang)

        if page_text and page_text.strip():
            all_text.append(f"[第{i+1}页]\n{page_text.strip()}")
            print(f"    第{i+1}页: 识别到 {len(page_text.strip())} 字符")
        else:
            print(f"    第{i+1}页: 无文本内容")

    pdf.close()

    full_text = "\n\n".join(all_text)
    print(f"  OCR完成，总计 {len(full_text)} 字符")
    return full_text


def call_dify_chat(query: str, inputs: Dict = None) -> str:
    """
    调用Dify Chat API

    Args:
        query: 用户输入的问题/内容
        inputs: 可选的输入变量字典

    Returns:
        LLM返回的内容
    """
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "query": query,
        "user": USER_ID,
        "response_mode": "blocking",  # 阻塞模式，等待完整响应
        "conversation_id": "",  # 空字符串表示新对话
        "inputs": inputs or {}
    }

    try:
        response = requests.post(
            f"{BASE_URL}/chat-messages",
            headers=headers,
            json=payload,
            timeout=300
        )
        response.raise_for_status()
        result = response.json()
        # Dify返回格式：{"answer": "...", "message_id": "...", ...}
        return result.get("answer", "")
    except Exception as e:
        print(f"Dify API调用失败: {e}")
        if hasattr(e, 'response') and e.response:
            print(f"  响应内容: {e.response.text[:500]}")
        return ""


# def call_dify_workflow(inputs: Dict) -> str:
#     """
#     调用Dify Workflow API

#     Args:
#         inputs: 工作流输入变量字典

#     Returns:
#         工作流执行结果
#     """
#     headers = {
#         "Authorization": f"Bearer {API_KEY}",
#         "Content-Type": "application/json"
#     }

#     payload = {
#         "inputs": inputs,
#         "user": USER_ID,
#         "response_mode": "blocking"
#     }

#     try:
#         response = requests.post(
#             f"{BASE_URL}/workflows/run",
#             headers=headers,
#             json=payload,
#             timeout=300
#         )
#         response.raise_for_status()
#         result = response.json()
#         # Dify Workflow返回格式：{"outputs": {...}, "status": "succeeded", ...}
#         return result.get("outputs", {}).get("text", "")
#     except Exception as e:
#         print(f"Dify Workflow API调用失败: {e}")
#         if hasattr(e, 'response') and e.response:
#             print(f"  响应内容: {e.response.text[:500]}")
#         return ""


def call_llm(prompt: str, content: str) -> str:
    """
    调用LLM处理文本内容
    使用Dify Chat API
    """
    full_query = f"{prompt}\n\n以下是SOC1报告的完整内容:\n{content}"
    return call_dify_chat(full_query)


def call_llm_with_images(prompt: str, images: List[str]) -> str:
    """
    调用LLM处理图像内容（扫描件PDF）
    使用Dify Chat API的files参数上传图像

    Dify Chatbot API文件上传流程:
    1. 先调用 /files/upload 上传文件，获取file_id
    2. 再调用 /chat-messages，在files参数中传入file_id列表
    """
    print(f"  上传 {len(images)} 张PDF页面图像...")

    try:
        headers = {
            "Authorization": f"Bearer {API_KEY}"
        }

        # Step 1: 上传每张图像，获取file_id
        file_ids = []
        for i, img_base64 in enumerate(images):
            # 将base64解码为二进制数据
            img_bytes = base64.b64decode(img_base64)

            # 使用multipart/form-data上传文件
            files_payload = {
                'file': (f'page_{i+1}.png', img_bytes, 'image/png'),
                'user': (None, USER_ID)
            }

            upload_response = requests.post(
                f"{BASE_URL}/files/upload",
                headers=headers,
                files=files_payload,
                timeout=60
            )
            upload_response.raise_for_status()
            upload_result = upload_response.json()
            file_id = upload_result.get("id")
            if file_id:
                file_ids.append(file_id)
                print(f"    已上传第{i+1}页图像，file_id: {file_id}")
            else:
                print(f"    [警告] 第{i+1}页上传失败")

        if not file_ids:
            print("  [错误] 所有图像上传失败")
            return ""

        # Step 2: 调用Chat API，传入file_ids
        print(f"  发送包含 {len(file_ids)} 个文件的请求到Dify...")
        chat_headers = {
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json"
        }

        chat_payload = {
            "query": prompt,
            "user": USER_ID,
            "response_mode": "blocking",
            "conversation_id": "",
            "inputs": {},
            "files": [
                {"type": "image", "id": file_id} for file_id in file_ids
            ]
        }

        chat_response = requests.post(
            f"{BASE_URL}/chat-messages",
            headers=chat_headers,
            json=chat_payload,
            timeout=600
        )
        chat_response.raise_for_status()
        result = chat_response.json()
        return result.get("answer", "")

    except Exception as e:
        print(f"Dify API调用失败（图像模式）: {e}")
        if hasattr(e, 'response') and e.response:
            print(f"  响应内容: {e.response.text[:500]}")
        return ""


def normalize_keys(data: Dict) -> Dict:
    """规范化JSON字段名，去除多余空格（只处理顶层键）"""
    if isinstance(data, dict):
        normalized = {}
        for key, value in data.items():
            # 只去除顶层key中的空格（如"CUEC 表格" -> "CUEC表格"）
            clean_key = key.replace(" ", "")
            # 递归处理嵌套结构，但保留嵌套dict中的原始键名
            if isinstance(value, dict):
                # 对于嵌套dict（如基本信息），保留原始键名
                normalized[clean_key] = value
            elif isinstance(value, list):
                # 对于列表中的dict项（如CUEC表格条目），保留原始键名
                normalized[clean_key] = value
            else:
                normalized[clean_key] = value
        return normalized
    return data


def parse_llm_response(response: str) -> Dict:
    """解析LLM返回的JSON"""
    try:
        if "{" in response and "}" in response:
            start = response.find("{")
            end = response.rfind("}") + 1
            json_str = response[start:end]
            raw_result = json.loads(json_str)
            # 规范化字段名
            return normalize_keys(raw_result)
    except json.JSONDecodeError as e:
        print(f"JSON解析失败: {e}")
    return {}


def process_single_pdf(pdf_path: str) -> Dict:
    """处理单个PDF文件"""
    filename = Path(pdf_path).name
    print(f"\n处理: {filename}")

    # 检测是否是扫描件
    is_scanned = is_scanned_pdf(pdf_path)

    # 构建prompt
    prompt = """你是一个SOC107报告内容提取专家。请从以下SOC107报告中提取三类内容，以JSON格式返回。

【提取要求】

**第一部分：基本信息**
- SOC报告名称：报告的完整标题
- 审计期间：例如 "October 1, 2024 to September 30, 2025"
- 审计师：例如 "KPMG LLP"、"PricewaterhouseCoopers"、"Ernst & Young"、"Deloitte"等

**第二部分：控制表格**

请严格按照PDF中原始表格的结构提取，不要拆分或合并表格行。

1. Complementary User Entity Controls (CUECs) 表格
   - 原始表格格式：两列，分别是"Control Objective"和"Customer's Responsibilities"
   - 每个Control Objective对应一行，其CUEC内容可能包含多段文字，请保持完整，不要拆分成多行
   - 例如：如果"Invoice Payment Processing"的CUEC包含4段"Customers are responsible for..."，这些应该合并在一个单元格内，保持原始表格结构

2. Complementary Subservice Organization Controls (CSOCs) 表格（如果存在）
   - 原始表格格式：两列，分别是"Impacted Control Objectives"和"Controls Expected to be Implemented at Subservice Organizations"
   - 每个Control Objective对应一行，保持完整内容

**第三部分：IT流程页码范围**
提取以下两个控制领域的页码范围，每个领域需要提取Section III和Section IV两个部分：

1. Change Management
   - Section III中的页码范围（描述部分）
   - Section IV中的页码范围（控制目标和测试结果部分）

2. Access Management（可能也叫User and Access Management或Identity and Access Management）
   - Section III中的页码范围
   - Section IV中的页码范围

【返回格式】

请严格按照以下JSON格式返回：

```json
{
    "基本信息": {
        "报告名称": "...",
        "审计期间": "...",
        "审计师": "..."
    },
    "CUEC表格": [
        {
            "Control Objective": "Change Management",
            "CUEC": "Customers are responsible for communicating with SAP Ariba regarding the timing and implementation of changes to their systems."
        },
        {
            "Control Objective": "Invoice Payment Processing",
            "CUEC": "Customers are responsible for configuring custom business rules for invoice processing reconciliation... Customers are responsible for controls over managing their own access control systems... Customers are responsible for controls over making additions... （完整内容，保持原始表格行结构）"
        }
    ],
    "CSOC表格": [
        {
            "Impacted Control Objectives": "Physical Security",
            "CSOC": "The colocation Data Centers are responsible for implementing physical security and environmental safeguards..."
        }
    ],
    "Change Management页码": {
        "Section III": ["Page 22", "Page 28"],
        "Section IV": ["Page 45 ~ Page 48"]
    },
    "Access Management页码": {
        "Section III": ["Page 23", "Page 29"],
        "Section IV": ["Page 63 ~ Page 67"]
    }
}
```

重要提示：
- CUEC表格：每个Control Objective只占一行，其对应的CUEC内容应完整保留（包括所有"Customers are responsible for..."段落），用换行符分隔各段落
- 如果报告中没有CSOC表格，CSOC表格返回空数组[]
- 页码格式：连续页用"Page X ~ Page Y"，单页用"Page X"
- 只返回JSON，不要有其他文字"""

    # 根据是否是扫描件选择处理方式
    if is_scanned:
        # 扫描件：使用OCR识别后发送文本给LLM
        print("  OCR识别扫描件PDF...")
        ocr_text = ocr_pdf_pages(pdf_path, resolution=200, lang='eng')
        print(f"  OCR文本长度: {len(ocr_text)} 字符")

        print("  调用LLM提取内容（OCR文本模式）...")
        response = call_llm(prompt, ocr_text)
    else:
        # 文本型PDF：使用文本模式
        print("  提取PDF文本...")
        pdf_text = extract_pdf_text(pdf_path)
        print(f"  文本长度: {len(pdf_text)} 字符")

        print("  调用LLM提取内容（文本模式）...")
        response = call_llm(prompt, pdf_text)

    # 解析结果
    print("  解析返回结果...")
    print(f"  LLM返回长度: {len(response)} 字符")
    result = parse_llm_response(response)
    print(f"  解析后的键: {list(result.keys())}")

    if not result:
        print("  [警告] 解析失败，返回空结果")
        result = {
            "基本信息": {"报告名称": "", "审计期间": "", "审计师": ""},
            "CUEC表格": [],
            "CSOC表格": [],
            "Change Management页码": {"Section III": [], "Section IV": []},
            "Access Management页码": {"Section III": [], "Section IV": []}
        }

    # 显示提取结果摘要
    basic = result.get("基本信息", {})
    print(f"  报告名称: {basic.get('报告名称', '')[:50]}...")
    print(f"  审计期间: {basic.get('审计期间', '')}")
    print(f"  审计师: {basic.get('审计师', '')}")
    print(f"  CUEC条数: {len(result.get('CUEC表格', []))}")
    print(f"  CSOC条数: {len(result.get('CSOC表格', []))}")
    print(f"  CM页码: {result.get('ChangeManagement页码', result.get('Change Management页码', {}))}")
    print(f"  AM页码: {result.get('AccessManagement页码', result.get('Access Management页码', {}))}")

    return {"文件名": filename, **result}


def write_to_excel(results: List[Dict], output_path: str):
    """将结果写入Excel，每个PDF单独一个Sheet"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for r in results:
            # Sheet名称（去掉.pdf后缀，限制长度）
            sheet_name = Path(r["文件名"]).stem[:30]

            # 构建行数据
            rows = []

            # ===== Part 1: 基本信息 =====
            rows.append({"类型": "1 基本信息", "字段": "", "内容": ""})
            basic = r.get("基本信息", {})
            rows.append({"类型": "", "字段": "SOC报告名称", "内容": basic.get("报告名称", "")})
            rows.append({"类型": "", "字段": "审计期间", "内容": basic.get("审计期间", "")})
            rows.append({"类型": "", "字段": "审计师", "内容": basic.get("审计师", "")})
            rows.append({"类型": "", "字段": "", "内容": ""})

            # ===== Part 2: 控制表格 =====
            rows.append({"类型": "2 控制表格输出", "字段": "", "内容": ""})
            rows.append({"类型": "2.1 Complementary User Entity Controls（CUECs）", "字段": "", "内容": ""})

            cuec = r.get("CUEC表格", [])
            if cuec:
                rows.append({"类型": "", "字段": "Control Objective", "内容": "CUEC"})
                for item in cuec:
                    rows.append({
                        "类型": "",
                        "字段": item.get("Control Objective", ""),
                        "内容": item.get("CUEC", "")
                    })
            else:
                rows.append({"类型": "", "字段": "（无CUEC内容）", "内容": ""})

            rows.append({"类型": "", "字段": "", "内容": ""})

            rows.append({"类型": "2.2 Complementary Subservice Organization Controls（CSOCs）", "字段": "", "内容": ""})

            csoc = r.get("CSOC表格", [])
            if csoc:
                rows.append({"类型": "", "字段": "Impacted Control Objectives", "内容": "CSOC"})
                for item in csoc:
                    rows.append({
                        "类型": "",
                        "字段": item.get("Impacted Control Objectives", ""),
                        "内容": item.get("CSOC", "")
                    })
            else:
                rows.append({"类型": "", "字段": "（该报告无CSOC内容）", "内容": ""})

            rows.append({"类型": "", "字段": "", "内容": ""})

            # ===== Part 3: 页码范围 =====
            rows.append({"类型": "3 IT流程页码输出", "字段": "", "内容": ""})
            rows.append({"类型": "", "字段": "Domain", "内容": "Page Range"})

            # 获取页码信息，兼容两种键名
            cm_pages = r.get("ChangeManagement页码") or r.get("Change Management页码", {})
            am_pages = r.get("AccessManagement页码") or r.get("Access Management页码", {})

            rows.append({"类型": "", "字段": "Change Management", "内容": ""})
            for section, pages in cm_pages.items():
                # 兼容SectionIII和Section III两种格式
                section_display = section.replace("Section", "Section ") if "Section" in section and " " not in section else section
                for page_range in pages:
                    rows.append({"类型": "", "字段": f"  {section_display}", "内容": page_range})

            rows.append({"类型": "", "字段": "Access Management", "内容": ""})
            for section, pages in am_pages.items():
                section_display = section.replace("Section", "Section ") if "Section" in section and " " not in section else section
                for page_range in pages:
                    rows.append({"类型": "", "字段": f"  {section_display}", "内容": page_range})

            # 写入Sheet
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n结果已保存到: {output_path}")


def main():
    """主函数"""
    # 从控制台输入文件夹路径
    data_dir = input("请输入PDF文件夹路径: ").strip()

    # 去除可能的引号
    if data_dir.startswith('"') and data_dir.endswith('"'):
        data_dir = data_dir[1:-1]
    if data_dir.startswith("'") and data_dir.endswith("'"):
        data_dir = data_dir[1:-1]

    # 检查路径是否存在
    if not os.path.exists(data_dir):
        print(f"错误: 路径不存在 - {data_dir}")
        return

    if not os.path.isdir(data_dir):
        print(f"错误: 输入的不是文件夹 - {data_dir}")
        return

    pdf_files = [f for f in os.listdir(data_dir) if f.endswith('.pdf')]

    if not pdf_files:
        print(f"错误: 该文件夹中没有PDF文件 - {data_dir}")
        return

    pdf_paths = [os.path.join(data_dir, f) for f in pdf_files]

    print("=" * 60)
    print("SOC1 Report Content Extractor (Dify API Version)")
    print(f"Dify URL: {BASE_URL}")
    print(f"API Key: {API_KEY[:20]}...")
    print(f"User ID: {USER_ID}")
    print("=" * 60)
    print(f"\n输入文件夹: {data_dir}")
    print(f"发现 {len(pdf_files)} 个PDF文件:")
    for f in pdf_files:
        print(f"  - {f}")

    results = []
    for pdf_path in pdf_paths:
        try:
            result = process_single_pdf(pdf_path)
            results.append(result)
        except Exception as e:
            print(f"\n错误: 处理 {pdf_path} 时出错: {e}")
            results.append({
                "文件名": Path(pdf_path).name,
                "错误": str(e),
                "基本信息": {"报告名称": "", "审计期间": "", "审计师": ""},
                "CUEC表格": [],
                "CSOC表格": [],
                "Change Management页码": {"Section III": [], "Section IV": []},
                "Access Management页码": {"Section III": [], "Section IV": []}
            })

    # 写入Excel - 输出到当前工作目录下的output文件夹
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "提取结果_Dify.xlsx")
    write_to_excel(results, output_path)

    print("\n" + "=" * 60)
    print("处理完成!")
    print(f"结果已保存到: {output_path}")
    print("=" * 60)

    # Windows打包后暂停，防止窗口自动关闭
    if sys.platform == 'win32':
        input("\n按回车键退出...")


if __name__ == "__main__":
    main()