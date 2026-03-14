import os
import shutil
import tempfile
from excel_tool import merge_excel_files, split_excel_by_columns
from openpyxl import Workbook

def create_test_excel(file_path, sheet_data):
    """
    创建测试用的Excel文件
    :param file_path: 文件路径
    :param sheet_data: 字典，键为sheet名称，值为数据列表
    """
    wb = Workbook()
    wb.remove(wb.active)
    
    for sheet_name, data in sheet_data.items():
        ws = wb.create_sheet(title=sheet_name)
        for row in data:
            ws.append(row)
    
    wb.save(file_path)
    wb.close()

def test_merge_excel():
    """
    测试合并Excel功能
    """
    print("测试合并Excel功能...")
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    
    try:
        # 创建测试文件1
        file1 = os.path.join(temp_dir, "file1.xlsx")
        create_test_excel(file1, {
            "Sheet1": [["姓名", "年龄", "城市"], ["张三", 25, "北京"], ["李四", 30, "上海"]],
            "Sheet2": [["产品", "价格"], ["A", 100], ["B", 200]]
        })
        
        # 创建测试文件2
        file2 = os.path.join(temp_dir, "file2.xlsx")
        create_test_excel(file2, {
            "Sheet1": [["姓名", "年龄", "城市"], ["王五", 35, "广州"], ["赵六", 40, "深圳"]],
            "Sheet3": [["部门", "人数"], ["技术部", 10], ["市场部", 5]]
        })
        
        # 合并文件
        output_file = os.path.join(temp_dir, "merged.xlsx")
        merge_excel_files([file1, file2], output_file)
        
        print(f"合并测试完成，结果保存在: {output_file}")
        print("合并测试通过！")
        
    finally:
        # 清理临时文件
        shutil.rmtree(temp_dir)

def test_split_excel():
    """
    测试拆分Excel功能
    """
    print("\n测试拆分Excel功能...")
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    
    try:
        # 创建测试文件
        input_file = os.path.join(temp_dir, "data.xlsx")
        create_test_excel(input_file, {
            "Sheet1": [["姓名", "年龄", "城市"], 
                      ["张三", 25, "北京"], 
                      ["李四", 30, "上海"], 
                      ["王五", 35, "北京"], 
                      ["赵六", 40, "上海"]],
            "Sheet2": [["产品", "类别", "价格"], 
                      ["A", "电子", 100], 
                      ["B", "服装", 200], 
                      ["C", "电子", 300], 
                      ["D", "服装", 400]]
        })
        
        # 按城市列拆分（索引2）
        output_dir = os.path.join(temp_dir, "split_by_city")
        split_excel_by_columns(input_file, output_dir, [2])
        
        # 按类别列拆分（索引1）
        output_dir2 = os.path.join(temp_dir, "split_by_category")
        split_excel_by_columns(input_file, output_dir2, [1])
        
        print("拆分测试完成！")
        print(f"按城市拆分结果保存在: {output_dir}")
        print(f"按类别拆分结果保存在: {output_dir2}")
        
    finally:
        # 清理临时文件
        shutil.rmtree(temp_dir)

def main():
    """
    运行所有测试
    """
    print("开始测试Excel工具...")
    test_merge_excel()
    test_split_excel()
    print("\n所有测试完成！")

if __name__ == "__main__":
    main()