import os
import pandas as pd
import sys

# 输出文件路径固定
output_file = './output.xlsx'  # 输出文件路径

def process_csv_files(input_directory, output_file):
    # 获取所有 CSV 文件
    csv_files = [f for f in os.listdir(input_directory) if f.endswith('.csv')]

    # 用来存储每个 CSV 文件的数据，字典名为文件名去后缀
    file_data_dict = {}

    # 用来存储所有 CSV 文件第一个字段去重后的值
    all_first_column_values = set()

    # 读取所有 CSV 文件并存入字典
    for csv_file in csv_files:
        # 获取文件的路径
        csv_path = os.path.join(input_directory, csv_file)

        try:
            # 使用 UTF-8-SIG 编码读取 CSV 文件（处理 BOM）
            df = pd.read_csv(csv_path, encoding='utf-8-sig', on_bad_lines='skip')  # 跳过错误行

            # 使用文件名（去掉扩展名）作为字典的键
            file_name_without_extension = os.path.splitext(csv_file)[0]
            file_data_dict[file_name_without_extension] = df

            # 提取每个文件第一个字段（列）的去重值，并加入到 all_first_column_values 中
            if not df.empty:
                first_column_values = df.iloc[:, 0].dropna().unique()  # 获取第一个字段的去重值
                all_first_column_values.update(first_column_values)

            print(f"文件 '{csv_file}' 成功读取。")
        except Exception as e:
            print(f"读取文件 '{csv_file}' 失败，错误: {e}")

    # 创建一个 Excel Writer 对象
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 写入所有 CSV 文件的数据
        for file_name, df in file_data_dict.items():
            # 对每个 CSV 文件的数据进行去重
            df_deduplicated = df.drop_duplicates()

            # 将去重后的数据写入 Excel，每个文件的数据作为单独的工作表
            df_deduplicated.to_excel(writer, sheet_name=file_name, index=False)

        # 将所有文件第一个字段的去重值写入一个新的工作表
        first_column_df = pd.DataFrame(list(all_first_column_values), columns=['Unique First Column Values'])
        first_column_df.to_excel(writer, sheet_name='Unique First Column Values', index=False)

    print(f"所有数据已保存到 {output_file}")

if __name__ == '__main__':
    # 确保命令行提供了输入目录路径
    if len(sys.argv) != 2:
        print("请提供输入目录路径。")
        sys.exit(1)

    # 获取命令行传入的路径
    input_directory = sys.argv[1]

    # 调用处理函数
    process_csv_files(input_directory, output_file)
