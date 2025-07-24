import pandas as pd
import os
import json
from pathlib import Path
from openpyxl import load_workbook
from typing import List, Dict, Any, Union

class RobustExcelWriter:
    """
    健壮的Excel写入器，借鉴DataProcessor类的异常处理机制
    支持文件占用、权限错误等异常情况的处理
    """
    
    def __init__(self, save_dir: str = "save"):
        """
        初始化Excel写入器
        
        Args:
            save_dir: 保存目录，默认为"save"
        """
        self.save_dir = Path(save_dir)
        self.save_dir.mkdir(exist_ok=True)
        print(f"Excel写入器初始化完成，保存目录: {self.save_dir}")
    
    def safe_write_excel(self, data: Union[List[Dict], pd.DataFrame], filename: str, 
                        sheet_name: str = "Sheet1") -> str:
        """
        安全写入Excel文件，处理各种异常情况
        
        Args:
            data: 要写入的数据，可以是字典列表或DataFrame
            filename: 文件名
            sheet_name: 工作表名称
            
        Returns:
            str: 实际保存的文件路径
        """
        output_path = self.save_dir / filename
        
        # 转换数据为DataFrame
        if isinstance(data, list):
            df = pd.DataFrame(data)
        elif isinstance(data, pd.DataFrame):
            df = data
        else:
            raise ValueError("数据必须是字典列表或DataFrame")
        
        # 如果DataFrame为空，创建空的DataFrame
        if df.empty:
            print(f"⚠️ 数据为空，创建空的Excel文件: {filename}")
        
        try:
            # 尝试正常写入
            df.to_excel(output_path, index=False, sheet_name=sheet_name)
            print(f"✓ Excel文件已保存: {output_path}")
            return str(output_path)
            
        except PermissionError:
            # 文件被占用（如在Excel中打开）
            temp_path = self.save_dir / f"{Path(filename).stem}_temp{Path(filename).suffix}"
            df.to_excel(temp_path, index=False, sheet_name=sheet_name)
            print(f"⚠️ 原文件被占用，已保存到临时文件: {temp_path}")
            return str(temp_path)
            
        except Exception as e:
            # 其他错误，使用备份文件
            backup_path = self.save_dir / f"{Path(filename).stem}_backup{Path(filename).suffix}"
            try:
                df.to_excel(backup_path, index=False, sheet_name=sheet_name)
                print(f"❌ 保存失败 ({e})，已保存到备份文件: {backup_path}")
                return str(backup_path)
            except Exception as backup_error:
                print(f"❌ 备份保存也失败: {backup_error}")
                # 最后尝试保存为CSV
                csv_path = self.save_dir / f"{Path(filename).stem}_emergency.csv"
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                print(f"🆘 紧急保存为CSV: {csv_path}")
                return str(csv_path)
    
    def safe_write_multi_sheet_excel(self, data_dict: Dict[str, Union[List[Dict], pd.DataFrame]], 
                                   filename: str) -> str:
        """
        安全写入多工作表Excel文件
        
        Args:
            data_dict: 字典，键为工作表名，值为数据
            filename: 文件名
            
        Returns:
            str: 实际保存的文件路径
        """
        output_path = self.save_dir / filename
        
        # 转换所有数据为DataFrame
        df_dict = {}
        for sheet_name, data in data_dict.items():
            if isinstance(data, list):
                df_dict[sheet_name] = pd.DataFrame(data)
            elif isinstance(data, pd.DataFrame):
                df_dict[sheet_name] = data
            else:
                print(f"⚠️ 跳过无效数据类型的工作表: {sheet_name}")
                continue
        
        if not df_dict:
            print("❌ 没有有效的数据可写入")
            return ""
        
        try:
            # 尝试正常写入多工作表
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in df_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 调整列宽
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
                        
                        # 设置合适的列宽，最大不超过50
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            print(f"✓ 多工作表Excel文件已保存: {output_path}")
            return str(output_path)
            
        except PermissionError:
            # 文件被占用时的处理
            temp_path = self.save_dir / f"{Path(filename).stem}_temp{Path(filename).suffix}"
            try:
                with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                    for sheet_name, df in df_dict.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"⚠️ 原文件被占用，已保存到临时文件: {temp_path}")
                return str(temp_path)
            except Exception as temp_error:
                print(f"❌ 临时文件保存失败: {temp_error}")
                return self._emergency_save_multiple(df_dict, filename)
                
        except Exception as e:
            print(f"❌ 多工作表保存失败: {e}")
            return self._emergency_save_multiple(df_dict, filename)
    
    def _emergency_save_multiple(self, df_dict: Dict[str, pd.DataFrame], base_filename: str) -> str:
        """紧急保存多个工作表为单独的文件"""
        saved_files = []
        base_name = Path(base_filename).stem
        
        for sheet_name, df in df_dict.items():
            emergency_filename = f"{base_name}_{sheet_name}_emergency.xlsx"
            try:
                path = self.safe_write_excel(df, emergency_filename, sheet_name)
                saved_files.append(path)
            except Exception as e:
                print(f"❌ 紧急保存工作表 {sheet_name} 失败: {e}")
        
        if saved_files:
            print(f"🆘 已紧急保存为 {len(saved_files)} 个单独文件")
            return saved_files[0]  # 返回第一个文件路径
        else:
            print("❌ 所有紧急保存尝试都失败了")
            return ""
    
    def export_analysis_results(self, json_file_path: str) -> tuple:
        """
        从JSON文件导出分析结果到Excel
        
        Args:
            json_file_path: JSON文件路径
            
        Returns:
            tuple: (基础Excel路径, 详细Excel路径)
        """
        try:
            # 读取JSON文件
            with open(json_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not data:
                print("⚠️ JSON文件为空，跳过Excel导出")
                return "", ""
            
            # 准备基础数据
            basic_data = []
            tag_detail_data = []
            
            for idx, item in enumerate(data):
                # 基础数据
                basic_row = {
                    'ID': idx + 1,
                    '游戏名称': item.get('game_name', ''),
                    '评论内容': item.get('review_content', ''),
                    '内容摘要': item.get('summary', ''),
                    '优点总结': item.get('pros', ''),
                    '缺点总结': item.get('cons', ''),
                    '内容标签': item.get('tags', '')
                }
                
                # 处理标签属性
                tag_attributes = item.get('tag_attributes', {})
                if isinstance(tag_attributes, dict):
                    tag_attr_str = ', '.join([f"{k}:{v}" for k, v in tag_attributes.items()])
                    basic_row['标签属性分类'] = tag_attr_str
                else:
                    basic_row['标签属性分类'] = str(tag_attributes) if tag_attributes else ''
                
                basic_data.append(basic_row)
                
                # 标签详细数据
                tags = item.get('tags', '')
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
            
            # 生成文件名
            json_path = Path(json_file_path)
            basic_filename = f"{json_path.stem}_analysis_results.xlsx"
            detailed_filename = f"{json_path.stem}_detailed_analysis.xlsx"
            
            # 保存基础Excel
            basic_path = self.safe_write_excel(basic_data, basic_filename, "游戏评论分析结果")
            
            # 保存详细Excel（多工作表）
            detailed_data = {
                "分析结果总览": basic_data,
                "标签详细分析": tag_detail_data
            }
            detailed_path = self.safe_write_multi_sheet_excel(detailed_data, detailed_filename)
            
            return basic_path, detailed_path
            
        except Exception as e:
            print(f"❌ 导出分析结果时发生错误: {e}")
            return "", ""

# 全局实例
excel_writer = RobustExcelWriter()

def export_analysis_to_excel_robust(json_file_path: str) -> tuple:
    """
    健壮的Excel导出函数，替代原有的导出功能
    
    Args:
        json_file_path: JSON文件路径
        
    Returns:
        tuple: (基础Excel路径, 详细Excel路径)
    """
    return excel_writer.export_analysis_results(json_file_path)

def auto_export_on_interrupt_robust(json_file_path: str):
    """
    中断时的健壮Excel导出
    
    Args:
        json_file_path: JSON文件路径
    """
    try:
        if os.path.exists(json_file_path):
            print(f"\n🔄 检测到程序中断，正在导出当前进度的Excel文件...")
            basic_path, detailed_path = export_analysis_to_excel_robust(json_file_path)
            
            if basic_path:
                print(f"✓ 中断保存 - 基础Excel: {basic_path}")
            if detailed_path:
                print(f"✓ 中断保存 - 详细Excel: {detailed_path}")
                
        else:
            print(f"⚠️ JSON文件 {json_file_path} 不存在，跳过Excel导出")
            
    except Exception as e:
        print(f"❌ 中断保存Excel文件时出现错误: {e}")

if __name__ == "__main__":
    # 测试代码
    test_data = [
        {
            'game_name': '测试游戏',
            'review_content': '这是一个测试评论',
            'summary': '测试摘要',
            'pros': '测试优点',
            'cons': '测试缺点',
            'tags': '测试标签1,测试标签2',
            'tag_attributes': {'测试标签1': '测试属性1', '测试标签2': '测试属性2'}
        }
    ]
    
    writer = RobustExcelWriter()
    writer.safe_write_excel(test_data, "test_output.xlsx")
    print("测试完成")
