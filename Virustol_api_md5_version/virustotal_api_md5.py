# -*- coding: utf-8 -*-
import requests
import pandas as pd

api_key = '替换为您的 VirusTotal API 密钥'

def check_md5(md5_hash):
    url = f'https://www.virustotal.com/api/v3/files/{md5_hash}'
    headers = {
        'x-apikey': api_key
    }
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            result = response.json()
            return result
        else:
            print(f"Error {response.status_code}: Unable to fetch data for {md5_hash}")
            return None
    except Exception as e:
        print(f"Exception occurred while fetching data for {md5_hash}: {e}")
        return None

def read_md5_from_file(file_path):
    """Read MD5 hashes from a file (supports .txt and .csv/.xlsx)"""
    md5_hashes = []
    
    if file_path.endswith('.txt'):
        with open(file_path, 'r') as file:
            md5_hashes = [line.strip() for line in file]
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
        md5_hashes = df.iloc[:, 0].dropna().tolist()
    elif file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
        md5_hashes = df.iloc[:, 0].dropna().tolist()
    else:
        print("Unsupported file format")
    
    return md5_hashes

def output_result(result):
    """Output the result based on VirusTotal's report"""
    if not result or 'data' not in result or 'attributes' not in result['data']:
        return "未找到"
    
    last_analysis_stats = result['data']['attributes'].get('last_analysis_stats', {})
    if sum(last_analysis_stats.values()) == last_analysis_stats.get('undetected', 0):
        return "白"
    else:
        return "是"

def main():
    # 示例：读取 MD5 值文件路径
    input_file_path = 'D:/Work/戎恒/word/path/virustotal_api_md5/MD5.xlsx'
    output_file_path = 'D:/Work/戎恒/word/path/virustotal_api_md5/MD5jg(副本).xlsx'
    
    # 读取所有 MD5 值
    md5_hashes = read_md5_from_file(input_file_path)
    
    # 创建 DataFrame 来保存结果
    results = []

    # 查询每个 MD5 值并在 VirusTotal 上获取结果
    for md5 in md5_hashes:
        result = check_md5(md5)
        vt_result = output_result(result)
        results.append([md5, vt_result])
        print(f"{md5}: {vt_result}")  # 打印结果以便实时查看进度

    # 将结果保存到新的 Excel 文件
    df_results = pd.DataFrame(results, columns=['MD5', '结果'])
    df_results.to_excel(output_file_path, index=False)

if __name__ == "__main__":
    main()