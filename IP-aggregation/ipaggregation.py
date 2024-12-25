import pandas as pd
import ipaddress
from typing import List, Dict
from collections import defaultdict

def split_ips(ip_cell: str) -> List[str]:
    """拆分单元格中的多个IP地址"""
    if pd.isna(ip_cell):
        return []
    # 处理可能的分隔符
    separators = [',', ';', ' ']
    ips = [ip_cell]
    for sep in separators:
        ips = [ip.strip() for item in ips for ip in item.split(sep) if ip.strip()]
    return [ip for ip in ips if is_valid_ip(ip)]

def is_valid_ip(ip: str) -> bool:
    """检查是否为有效的IP地址"""
    try:
        ipaddress.ip_address(ip.strip())
        return True
    except:
        return False

def get_c_class_network(ip: str) -> str:
    """获取IP地址的C段网络"""
    try:
        ip_parts = ip.strip().split('.')
        return f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.0/24"
    except:
        return None

def find_optimal_subnet(ip_list: List[str]) -> Dict:
    """为一组IP找到最优子网"""
    # 如果只有一个IP，返回单个IP的信息
    if len(ip_list) == 1:
        ip = ip_list[0].strip()
        return {
            'network': ip,
            'netmask': '255.255.255.255',
            'cidr': 32,
            'first_ip': ip,
            'last_ip': ip,
            'total_ips': 1,
            'included_ips': ip
        }
    
    # 多个IP的处理逻辑
    ip_list = [ip for ip in ip_list if pd.notna(ip) and ip.strip()]
    ips = [ipaddress.IPv4Address(ip.strip()) for ip in ip_list]
    min_ip = min(ips)
    max_ip = max(ips)
    
    ip_range = int(max_ip) - int(min_ip)
    host_bits = (ip_range).bit_length()
    host_bits = max(2, host_bits)
    prefix_length = 32 - host_bits
    
    network = ipaddress.IPv4Network(f"{min_ip}/{prefix_length}", strict=False)
    while not all(ip in network for ip in ips):
        prefix_length -= 1
        network = ipaddress.IPv4Network(f"{min_ip}/{prefix_length}", strict=False)
    
    return {
        'network': str(network.network_address),
        'netmask': str(network.netmask),
        'cidr': network.prefixlen,
        'first_ip': str(network.network_address + 1),
        'last_ip': str(network.broadcast_address - 1),
        'total_ips': network.num_addresses - 2,
        'included_ips': ', '.join(ip_list)
    }

def process_ip_file(input_file: str, output_file: str):
    try:
        df = pd.read_excel(input_file)
        
        # 处理所有IP地址
        all_ips = []
        for cell in df.iloc[:, 0]:
            all_ips.extend(split_ips(str(cell)))
        
        # 按C段分组
        c_class_groups = defaultdict(list)
        for ip in all_ips:
            if ip:  # 确保IP不为空
                c_class = get_c_class_network(ip)
                if c_class:
                    c_class_groups[c_class].append(ip)
        
        all_results = []
        
        # 对每个C段分别处理
        for c_class, ips in c_class_groups.items():
            # 如果C段只有一个IP，直接处理为单个IP
            if len(ips) == 1:
                subnet_info = find_optimal_subnet([ips[0]])
            else:
                subnet_info = find_optimal_subnet(ips)
                
            all_results.append({
                '网络地址': subnet_info['network'],
                '子网掩码': subnet_info['netmask'],
                'CIDR表示': f"/{subnet_info['cidr']}",
                '第一个可用IP': subnet_info['first_ip'],
                '最后一个可用IP': subnet_info['last_ip'],
                '可用IP数量': subnet_info['total_ips'],
                '包含的IP地址': subnet_info['included_ips']
            })
        
        result_df = pd.DataFrame(all_results)
        result_df.to_excel(output_file, index=False)
        
        print(f"处理完成，结果已保存到 {output_file}")
        print(f"处理的IP地址总数: {len(all_ips)}")
        return True
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        return False

if __name__ == "__main__":
    input_file = "IPinput.xlsx"
    output_file = "IPoutput.xlsx"
    process_ip_file(input_file, output_file)