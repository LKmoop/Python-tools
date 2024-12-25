import pandas as pd
import re

def extract_ips(ip_str):
    # 使用正则表达式匹配IP地址
    ip_pattern = r'\b(?:\d{1,3}\.){3}\d{1,3}\b'
    return re.findall(ip_pattern, str(ip_str))

def process_excel():
    # 读取Excel文件
    df = pd.read_excel('Test_input.xlsx')  # 请替换为输入文件名
    
    # 创建新的数据列表
    new_data = []
    
    # 遍历原始数据框的每一行
    for index, row in df.iterrows():
        network_address = row['网络地址']
        subnet_mask = row['子网掩码']
        ip_addresses = extract_ips(row['包含的IP地址'])
        
        # 对于每个提取出的IP地址创建一个新行
        for ip in ip_addresses:
            new_data.append({
                '网络地址': network_address,
                '子网掩码': subnet_mask,
                'IP地址': ip
            })
    
    # 创建新的数据框
    new_df = pd.DataFrame(new_data)
    
    # 保存到新的Excel文件
    new_df.to_excel('Test_output.xlsx', index=False)

if __name__ == '__main__':
    process_excel()
