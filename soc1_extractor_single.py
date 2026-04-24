#!/usr/bin/env python3
"""
SOC1 Report Content Extractor - Single Pass LLM Version
一次性将PDF内容发送给LLM，提取所有需要的内容
支持扫描件PDF（使用视觉模型处理图像）
"""

import pdfplumber
import pandas as pd
import os
import json
import base64
from pathlib import Path
from typing import Dict, List, Optional
import requests
from io import BytesIO

# API配置
API_KEY = "sk-sp-68b6ce43fb86444f9e8470622ac72baa"
BASE_URL = "https://coding.dashscope.aliyuncs.com/v1"

# 模型配置
MODEL = "qwen3.5-plus"  # 支持视觉的模型

# 扫描件检测阈值（前5页文本总量低于此值判定为扫描件）
SCAN_THRESHOLD = 500


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


def call_llm(prompt: str, content: str) -> str:
    """调用LLM API（文本模式）"""
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    full_prompt = f"{prompt}\n\n以下是SOC1报告的完整内容:\n{content}"

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": full_prompt}]
    }

    try:
        response = requests.post(
            f"{BASE_URL}/chat/completions",
            headers=headers,
            json=payload,
            timeout=300
        )
        response.raise_for_status()
        result = response.json()
        return result["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"LLM API调用失败: {e}")
        return ""


def call_llm_with_images(prompt: str, images: List[str]) -> str:
    """
    调用LLM API（图像模式）
    用于处理扫描件PDF
    """
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    # 构建多图像消息
    content_parts = []

    # 先添加所有图像
    for i, img_base64 in enumerate(images):
        content_parts.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{img_base64}"
            }
        })

    # 最后添加文本prompt
    content_parts.append({
        "type": "text",
        "text": prompt
    })

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": content_parts}]
    }

    try:
        print(f"  发送 {len(images)} 张图像到LLM...")
        response = requests.post(
            f"{BASE_URL}/chat/completions",
            headers=headers,
            json=payload,
            timeout=600  # 图像处理需要更长超时
        )
        response.raise_for_status()
        result = response.json()
        return result["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"LLM API调用失败（图像模式）: {e}")
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
    prompt = """你是一个SOC1报告内容提取专家。请从以下SOC1报告中提取三类内容，以JSON格式返回。

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
        # 扫描件：使用图像模式
        print("  提取PDF页面图像...")
        images = extract_all_pages_as_images(pdf_path, resolution=150)
        print(f"  共 {len(images)} 页图像")

        print("  调用LLM提取内容（图像模式）...")
        response = call_llm_with_images(prompt, images)
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
    data_dir = "/Users/julian/soc/data"
    pdf_files = [f for f in os.listdir(data_dir) if f.endswith('.pdf')]
    pdf_paths = [os.path.join(data_dir, f) for f in pdf_files]

    print("=" * 60)
    print("SOC1 Report Content Extractor (Single Pass LLM)")
    print(f"API: {BASE_URL}")
    print(f"Model: {MODEL}")
    print("=" * 60)
    print(f"\n发现 {len(pdf_files)} 个PDF文件:")
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

    # 写入Excel
    output_path = "/Users/julian/soc/output/提取结果.xlsx"
    os.makedirs("/Users/julian/soc/output", exist_ok=True)
    write_to_excel(results, output_path)

    print("\n" + "=" * 60)
    print("处理完成!")
    print("=" * 60)


if __name__ == "__main__":
    main()