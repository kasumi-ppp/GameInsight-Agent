import pandas as pd
import json
from typing import List, Dict, Any
from pathlib import Path

def export_analysis_to_excel(json_file_path: str, excel_output_path: str = None) -> str:
    """
    将JSON分析结果导出为Excel文件

    Args:
        json_file_path: JSON文件路径
        excel_output_path: Excel输出路径，如果为None则自动生成

    Returns:
        str: 生成的Excel文件路径
    """
    # 读取JSON文件
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 如果未指定输出路径，则自动生成
    if excel_output_path is None:
        json_path = Path(json_file_path)
        excel_output_path = json_path.parent / f"{json_path.stem}_analysis_results.xlsx"

    # 转换数据格式
    excel_data = []

    for item in data:
        # 基础字段
        row = {
            '游戏名称': item.get('game_name', ''),
            '评论内容': item.get('review_content', ''),
            '内容摘要': item.get('summary', ''),
            '优点总结': item.get('pros', ''),
            '缺点总结': item.get('cons', ''),
            '内容标签': item.get('tags', '')
        }

        # 处理标签属性分类
        tag_attributes = item.get('tag_attributes', {})
        if isinstance(tag_attributes, dict):
            # 将标签属性转换为字符串格式
            tag_attr_str = ', '.join([f"{k}:{v}" for k, v in tag_attributes.items()])
            row['标签属性分类'] = tag_attr_str
        else:
            row['标签属性分类'] = str(tag_attributes) if tag_attributes else ''

        excel_data.append(row)

    # 创建DataFrame
    df = pd.DataFrame(excel_data)

    # 导出到Excel
    with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='游戏评论分析结果', index=False)

        # 调整列宽
        worksheet = writer.sheets['游戏评论分析结果']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            # 设置合适的列宽，最大不超过50
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"Excel文件已导出到: {excel_output_path}")
    return str(excel_output_path)

def export_with_tag_breakdown(json_file_path: str, excel_output_path: str = None) -> str:
    """
    将JSON分析结果导出为Excel文件，包含标签详细分解

    Args:
        json_file_path: JSON文件路径
        excel_output_path: Excel输出路径，如果为None则自动生成
    
    Returns:
        str: 生成的Excel文件路径
    """
    # 读取JSON文件
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 如果未指定输出路径，则自动生成
    if excel_output_path is None:
        json_path = Path(json_file_path)
        excel_output_path = json_path.parent / f"{json_path.stem}_detailed_analysis.xlsx"
    
    # 主要数据表
    main_data = []
    # 标签详细表
    tag_detail_data = []
    
    for idx, item in enumerate(data):
        # 主表数据
        main_row = {
            'ID': idx + 1,
            '游戏名称': item.get('game_name', ''),
            '评论内容': item.get('review_content', ''),
            '内容摘要': item.get('summary', ''),
            '优点总结': item.get('pros', ''),
            '缺点总结': item.get('cons', ''),
            '内容标签': item.get('tags', '')
        }
        main_data.append(main_row)
        
        # 标签详细表数据
        tags = item.get('tags', '')
        tag_attributes = item.get('tag_attributes', {})
        
        if tags and isinstance(tags, str):
            tag_list = [tag.strip() for tag in tags.split(',') if tag.strip()]
            for tag in tag_list:
                tag_row = {
                    'ID': idx + 1,
                    '游戏名称': item.get('game_name', ''),
                    '标签名称': tag,
                    '标签属性': tag_attributes.get(tag, '') if isinstance(tag_attributes, dict) else ''
                }
                tag_detail_data.append(tag_row)
    
    # 创建DataFrame
    main_df = pd.DataFrame(main_data)
    tag_df = pd.DataFrame(tag_detail_data)
    
    # 导出到Excel（多个工作表）
    with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
        # 主要分析结果表
        main_df.to_excel(writer, sheet_name='分析结果总览', index=False)
        
        # 标签详细表
        if not tag_df.empty:
            tag_df.to_excel(writer, sheet_name='标签详细分析', index=False)
        
        # 调整列宽
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # 设置合适的列宽
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"详细Excel文件已导出到: {excel_output_path}")
    return str(excel_output_path)

if __name__ == "__main__":
    # 示例用法
    json_path = "analysis_results.json"
    
    # 基础导出
    export_analysis_to_excel(json_path)
    
    # 详细导出（包含标签分解）
    export_with_tag_breakdown(json_path)
