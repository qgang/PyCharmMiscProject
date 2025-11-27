import pandas as pd
import sys
from openpyxl.styles import Alignment

def read_excel_file(file_path):
    try:
        # 读取 Excel 文件
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print("错误: 文件未找到，请检查文件路径是否正确。")
    except Exception as e:
        print(f"错误: 发生了一个未知错误: {e}")
    return None

def handle_score(score):
    # 处理表头，删除第一行，且将第二行数据作为新表头
    score.iloc[1, :4] = score.iloc[0, :4]
    new_score = score[2:]
    new_score.columns = score.iloc[1]
    new_score = new_score.reset_index(drop=True) #重置索引

    #处理列
    target_str = '答案'
    columns_to_drop = [col for col in new_score.columns if target_str in col]
    new_score = new_score.drop(columns=columns_to_drop)
    new_score = new_score.drop(columns='学号')

    #重命名header
    rename_dict = {}
    for col in new_score.columns:
        idx = col.find('（')
        rename_dict[col] = col[:idx] if idx != -1 else col

    new_score = new_score.rename(columns=rename_dict)

    #排序
    return new_score.sort_values(by='考号')

def handle_point(score_sheet):
    # 定义列名和题型区域
    columns = ['','全卷','1卷','2卷','听力','语法填空','选词填空','完型填空','阅读','六选四','概要','翻译','作文']
    point_sheet = pd.DataFrame(columns=columns)
    point_sheet['']=score_sheet.columns[3:]

    # 全卷、1卷、2卷
    sections = {
        '全卷':(0, 1),
        '1卷':(1, 2),
        '2卷':(2, 3)
    }
    for sections, (row_idx, col_idx) in sections.items():
        point_sheet.iloc[row_idx, col_idx] = 1

    # 听力
    point_sheet.loc[3:22, '听力'] = 1
    return point_sheet

def save_to_excel(point_sheet, output_file):
    try:
        # 写入Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            point_sheet.to_excel(writer, sheet_name="考点划分表", index=False)

            # 获取工作簿和工作表对象
            wb = writer.book
            ws = writer.sheets["考点划分表"]

            # 冻结首行
            ws.freeze_panes = "A2"

            for row in ws.iter_rows(min_row=1, min_col=1):  # 从第2行第2列开始
                for cell in row:
                    # 设置单元格居中对齐
                    cell.alignment = Alignment(horizontal='center', vertical='center')
    except Exception as e:
        print(f"保存Excel失败: {e}")
        raise

if __name__ == "__main__":
    input_file = './score.xls'
    output_file = './point.xlsx'
    try:
        print("读取分数表......")
        score = read_excel_file(input_file)
        if score is None:
            print("分数表读取失败")
            sys.exit(1)
        score_sheet = handle_score(score)

        print("生成考点划分表......")
        point_sheet = handle_point(score_sheet)

        save_to_excel(point_sheet, output_file)
    except Exception as e:
        print(f"错误: 发生了一个未知错误:{e}")
        sys.exit(1)

    print("考点划分初表生成结束！")
