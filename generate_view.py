import pandas as pd

# 读取Excel文件中的两个工作表
category_path = "C:\\Users\\lpy16\\Desktop\\Category.xlsx"
product_path = "C:\\Users\\lpy16\\Desktop\\Product.xlsx"
category_df = pd.read_excel(category_path, sheet_name='category')
product_df = pd.read_excel(product_path, sheet_name='product')

# 创建输出文件
output_file_path = "C:\\Users\\lpy16\\Desktop\\output.txt"

with open(output_file_path, 'w', encoding='utf-8') as output_file:
    # 遍历category表中的每一行
    for _, category_row in category_df.iterrows():
        category_id = category_row['ForeignId']
        category_name = category_row['Name']

        # 写入类别Id和name
        output_file.write(f"{category_id}: {category_name}\n")

        # 找到与当前类别Id匹配的所有产品
        matching_products = product_df[product_df['CategoryForeignId'] == category_id]

        # 写入匹配的每个产品的name和code
        for _, product_row in matching_products.iterrows():
            product_name = product_row['Name']
            product_code = product_row.get('Code', None)
            if pd.isna(product_code) or product_code == "":
                # 如果code为空，打印name
                output_file.write(f"- {product_name}\n")
            else:
                # 如果code不为空，打印name和code
                output_file.write(f"- {product_name} ({product_code})\n")
        
        # 添加一个空行来分隔不同的类别
        output_file.write("\n")

print(f"输出文件已生成：{output_file_path}")
