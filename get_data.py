file_path = "附件2：2024.1-8生产系统数据 4.14.xlsx"

sheet_names = [
    "汇总",
    "硅料单耗计算",
    "耗材价格",
    "销售收入",
    "生产变动成本",
    "生产公用成本",
    "人工成本",
    "销售费用",
    "管理费用",
    "财务费用"
]

import pandas as pd

def get_table1(row_number: int):
    if not (2 <= row_number <= 14) or row_number == 11:
        raise ValueError("行号必须在 2-14 且不能为 11")
    sheet_name = sheet_names[0] 
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="B:I")
        row_data = df.iloc[row_number - 2].tolist()
        return row_data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")
    
def get_table2(row: int, column: str):
    if not (3 <= row <= 6):
        raise ValueError("行号必须在 3-6 范围内")
    sheet_name = sheet_names[1]
    col_index = ord(column) - ord('A')
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=None)
        data = [df.iloc[row - 2 + i * 8][col_index] for i in range(8)]
        return data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")

def get_table3(row_number: int):
    if not 3 <= row_number <= 25:
        raise ValueError("行号错误")
    sheet_name = sheet_names[2] 
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="C:J")
        row_data = df.iloc[row_number - 2].tolist()
        return row_data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")
    
def get_table4(row: int, column: str):
    # 检查输入的行号是否在3到7这个规定的范围内
    if not (3 <= row <= 7):
        # 如果行号不在规定范围内，抛出ValueError异常，并给出错误提示
        raise ValueError("行号必须在 3 - 7 范围内")
    # 选择工作表名称列表中的第四个工作表（索引为3）
    sheet_name = sheet_names[3]
    # 将输入的列名（字母）转换为对应的列索引，通过ASCII码值相减得到
    col_index = ord(column) - ord('A')
    try:
        # 读取指定路径的Excel文件中选定的工作表，并只读取B列到F列的数据
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=None)
        # 这里存在一个错误，row_number未定义，推测应该是row，该行代码逻辑上可能多余，暂且保留
        # 通过列表推导式，按每间隔7行的方式提取指定列的数据，i从0到7
        data = [df.iloc[row-2 + i * 7, col_index] for i in range(8)]
        # 返回提取到的数据列表
        return data
    except Exception as e:
        # 如果在读取Excel文件或提取数据过程中出现异常，抛出RuntimeError异常，并给出错误提示
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")
    
def get_table5(row: int, column: str):
    if not (3 <= row <= 9):
        raise ValueError("行号必须在 3-31 范围内")
    sheet_name = sheet_names[4]
    col_index = ord(column) - ord('A')
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=None)
        data = [df.iloc[row - 2 + i * 32][col_index] for i in range(8)]
        return data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")
    
def get_table6(row: int, column: str):
    if not (3 <= row <= 9):
        raise ValueError("行号必须在 3-9 范围内")
    sheet_name = sheet_names[5]
    col_index = ord(column) - ord('A')
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=None)
        data = [df.iloc[row - 2 + i * 10][col_index] for i in range(8)]
        return data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")

def get_table8(row_number: int):
    if not 3 <= row_number <= 20:
        raise ValueError("行号错误")
    sheet_name = sheet_names[7] 
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="B:I")
        row_data = df.iloc[row_number - 2].tolist()
        return row_data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")
    
def get_table9(row_number: int):
    if not 3 <= row_number <= 24:
        raise ValueError("行号错误")
    sheet_name = sheet_names[8] 
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="B:I")
        row_data = df.iloc[row_number - 2].tolist()
        return row_data
    except Exception as e:
        raise RuntimeError(f"读取 Excel 文件时出错: {e}")

def plot_ele_wat():#电镀金刚线+水性切割液
    ele = get_table3(10)
    wat = get_table3(11)
    months = np.array(range(1, 9))

    # 分别用自定义的指数函数拟合电镀金刚线和水性切割液的数据
    ele_fit = fit_custom_function(months, ele, exponential_function)
    wat_fit = fit_custom_function(months, wat, exponential_function)

    # 获取电镀金刚线拟合函数的参数
    ele_params = curve_fit(exponential_function, months, ele)[0]
    ele_function_str = f"y = {ele_params[0]:.4f} * exp(-{ele_params[1]:.4f} * x) + {ele_params[2]:.4f}"
    print(f"电镀金刚线拟合函数: {ele_function_str}")

    # 获取水性切割液拟合函数的参数
    wat_params = curve_fit(exponential_function, months, wat)[0]
    wat_function_str = f"y = {wat_params[0]:.4f} * exp(-{wat_params[1]:.4f} * x) + {wat_params[2]:.4f}"
    print(f"水性切割液拟合函数: {wat_function_str}")

    # 计算电镀金刚线的CV系数
    ele_residuals = np.array(ele) - ele_fit(months)
    ele_cv = (np.std(ele_residuals) / np.mean(ele)) * 100
    print(f"电镀金刚线的CV系数: {ele_cv}%")

    # 计算水性切割液的CV系数
    wat_residuals = np.array(wat) - wat_fit(months)
    wat_cv = (np.std(wat_residuals) / np.mean(wat)) * 100
    print(f"水性切割液的CV系数: {wat_cv}%")

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))

    # 绘制电镀金刚线的图
    ax1.plot(months, ele, label='电镀金刚线')
    ax1.plot(months, ele_fit(months), linestyle='--', label='电镀金刚线拟合曲线')
    ax1.set_xlabel('月份')
    ax1.set_ylabel('KM/万片')
    ax1.set_title('电镀金刚线数据拟合')
    ax1.legend()
    ax1.grid(True)

    # 绘制水性切割液的图
    ax2.plot(months, wat, label='水性切割液')
    ax2.plot(months, wat_fit(months), linestyle='--', label='水性切割液拟合曲线')
    ax2.set_xlabel('月份')
    ax2.set_ylabel('KG/万片')
    ax2.set_title('水性切割液数据拟合')
    ax2.legend()
    ax2.grid(True)

    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    # # 测试 get_table1 函数
    try:
        row_number = 6
        result_table1 = get_table1(row_number)
        print(f"get_table1({row_number}) 返回结果: {result_table1}")
    except Exception as e:
        print(f"测试 get_table1 时出错: {e}")

    # 测试 get_table2 函数
    try:
        row = 4
        column = "H"
        result_table2 = get_table2(row, column)
        print(f"get_table2({row}, '{column}') 返回结果: {result_table2}")
    except Exception as e:
        print(f"测试 get_table2 时出错: {e}")
