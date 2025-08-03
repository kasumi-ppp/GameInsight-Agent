import os
import glob
from pathlib import Path
from typing import List, Optional


def list_data_files(data_dir: str = "data") -> List[str]:
    """
    列出data目录下的所有支持的数据文件
    
    Args:
        data_dir: 数据文件目录路径
    
    Returns:
        数据文件路径列表
    """
    if not os.path.exists(data_dir):
        return []
    
    # 支持的文件格式
    supported_extensions = ['*.xlsx', '*.xls', '*.csv', '*.txt']
    
    data_files = []
    for ext in supported_extensions:
        pattern = os.path.join(data_dir, ext)
        data_files.extend(glob.glob(pattern))
    
    # 按文件名排序
    data_files.sort()
    return data_files


def select_data_file(data_dir: str = "data") -> Optional[str]:
    """
    交互式选择数据文件
    
    Args:
        data_dir: 数据文件目录路径
    
    Returns:
        选中的文件路径，如果取消则返回None
    """
    data_files = list_data_files(data_dir)
    
    if not data_files:
        print(f"❌ 在 '{data_dir}' 目录下未找到任何支持的数据文件")
        print("支持的格式：.xlsx, .xls, .csv, .txt")
        return None
    
    print(f"\n📁 在 '{data_dir}' 目录下找到以下数据文件：")
    print("=" * 50)
    
    for i, file_path in enumerate(data_files, 1):
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        file_size_mb = file_size / (1024 * 1024)
        
        print(f"{i:2d}. {file_name}")
        print(f"     路径: {file_path}")
        print(f"     大小: {file_size_mb:.2f} MB")
        print()
    
    print("=" * 50)
    
    while True:
        try:
            choice = input(f"请选择要分析的文件 (1-{len(data_files)}) 或输入 'q' 退出: ").strip()
            
            if choice.lower() == 'q':
                print("👋 已取消文件选择")
                return None
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(data_files):
                selected_file = data_files[choice_num - 1]
                print(f"✅ 已选择文件: {os.path.basename(selected_file)}")
                return selected_file
            else:
                print(f"❌ 请输入 1-{len(data_files)} 之间的数字")
                
        except ValueError:
            print("❌ 请输入有效的数字或 'q'")
        except KeyboardInterrupt:
            print("\n👋 已取消文件选择")
            return None


def get_file_info(file_path: str) -> dict:
    """
    获取文件的基本信息
    
    Args:
        file_path: 文件路径
    
    Returns:
        包含文件信息的字典
    """
    if not os.path.exists(file_path):
        return {}
    
    file_stat = os.stat(file_path)
    file_info = {
        'name': os.path.basename(file_path),
        'path': file_path,
        'size_bytes': file_stat.st_size,
        'size_mb': file_stat.st_size / (1024 * 1024),
        'extension': Path(file_path).suffix.lower(),
        'modified_time': file_stat.st_mtime
    }
    
    return file_info


def validate_data_file(file_path: str) -> bool:
    """
    验证数据文件是否有效
    
    Args:
        file_path: 文件路径
    
    Returns:
        文件是否有效
    """
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在: {file_path}")
        return False
    
    if os.path.getsize(file_path) == 0:
        print(f"❌ 文件为空: {file_path}")
        return False
    
    extension = Path(file_path).suffix.lower()
    supported_extensions = ['.xlsx', '.xls', '.csv', '.txt']
    
    if extension not in supported_extensions:
        print(f"❌ 不支持的文件格式: {extension}")
        print(f"支持的格式: {', '.join(supported_extensions)}")
        return False
    
    return True
