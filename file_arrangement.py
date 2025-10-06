import os
import shutil
import typing as tp

# 扫描特定文件夹中后缀名为.xlsx和.xls文件
def scan_file(folder_path: str) -> tp.List[str]:
    """扫描特定文件夹中后缀名为.xlsx和.xls文件"""
    file_list = []
    if not os.path.exists(folder_path):
        print(f"文件夹 {folder_path} 不存在")
        return file_list
    
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
            file_list.append(file_path)
    return file_list

# 根据关键词和scan_file返回的文件列表，创建文件夹并移动文件
def classify_files(root: str, folder_name: str, keywords: tp.List[str], file_list: tp.List[str]) -> None:
    """根据关键词和scan_file返回的文件列表，创建文件夹并移动文件"""
    folder_path = os.path.join(root, folder_name)
    
    # 创建基础文件夹
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    # 创建关键词对应的子文件夹
    for keyword in keywords:
        keyword_folder = os.path.join(folder_path, keyword)
        if not os.path.exists(keyword_folder):
            os.makedirs(keyword_folder)
    
    # 创建'其他'文件夹
    other_folder = os.path.join(folder_path, '其他')
    if not os.path.exists(other_folder):
        os.makedirs(other_folder)
    
    for file_path in file_list:
        file_name = os.path.basename(file_path)
        moved = False
        for keyword in keywords:
            if keyword in file_name:
                target_path = os.path.join(folder_path, keyword, file_name)
                try:
                    shutil.move(file_path, target_path)
                    moved = True
                    print(f"已将 {file_name} 移动到 {keyword} 文件夹")
                    break
                except Exception as e:
                    print(f"移动 {file_name} 到 {keyword} 文件夹失败: {e}")
        if not moved:
            target_path = os.path.join(folder_path, '其他', file_name)
            try:
                shutil.move(file_path, target_path)
                print(f"已将 {file_name} 移动到 其他 文件夹")
            except Exception as e:
                print(f"移动 {file_name} 到 其他 文件夹失败: {e}")



class Keywords:
    """
    关键词容器
    set_keywords: 设置关键词
    get_keywords: 获取关键词
    add_keywords: 添加关键词
    save_keywords: 保存关键词
    load_keywords: 加载关键词
    """
    def __init__(self) -> None:
        self.keywords = []
        self.keywords_file_path = 'keywords.txt'
        self.isset = 0

    def set_keywords(self, keywords: tp.Union[tp.List[str], tp.Tuple[str]]) -> None:
        """设置关键词"""
        if self.isset == 0:
            self.keywords = list(keywords)
            self.isset = 1
        else:
            print('已设置过关键词，不能重复设置')

    def add_keywords(self, keywords: tp.Union[tp.List[str], tp.Tuple[str]]) -> None:
        """添加关键词"""
        if self.keywords != []:
            for keyword in list(keywords):
                if keyword not in self.keywords:
                    self.keywords.append(keyword)
        else:
            print('请先设置关键词')
        
    
    def save_keywords(self) -> None:
        """保存关键词"""
        if self.isset == 1:
            with open(self.keywords_file_path, 'w', encoding='utf-8') as f:
                for keyword in self.keywords:
                    f.write(keyword + '\n')
        else:
            print('请先设置关键词')

    def load_keywords(self) -> tp.List[str]:
        """加载关键词"""
        if os.path.exists(self.keywords_file_path):
            with open(self.keywords_file_path, 'r', encoding='utf-8') as f:
                return [line.strip() for line in f.readlines()]
        else:
            print('关键词文件不存在')
            return []



if __name__ == '__main__':
    root = 'D:'
    folder_name = 'test'
    keywords = Keywords()
    keywords.set_keywords(['1', '2', '3'])
    keywords.save_keywords()
    
    source_folder = os.path.join(root, folder_name)
    print(f"开始扫描文件夹: {source_folder}")
    file_list = scan_file(source_folder)
    
    if not file_list:
        print("没有找到.xlsx或.xls文件")
    else:
        print(f"找到 {len(file_list)} 个Excel文件")
        
        # 创建目标文件夹结构
        target_root = os.path.join(root, f"{folder_name}_分类结果")
        print(f"开始分类文件到: {target_root}")
        classify_files(target_root, '', keywords.load_keywords(), file_list)
        print("文件分类完成")
