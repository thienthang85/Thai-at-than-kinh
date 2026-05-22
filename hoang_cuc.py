import sys, io

# --- MAPPING 64 QUẺ THEO TRẬT TỰ VĂN VƯƠNG ---
# Mỗi quẻ được mã hóa bằng chuỗi nhị phân (từ dưới lên trên, Hào 1 đến Hào 6)
# 0: Âm, 1: Dương
WEN_WANG_64 = {
    1: {"name": "Thuần Kiền", "bin": "111111"},
    2: {"name": "Thuần Khôn", "bin": "000000"},
    3: {"name": "Thủy Lôi Truân", "bin": "100010"},
    4: {"name": "Sơn Thủy Mông", "bin": "010001"},
    5: {"name": "Thủy Thiên Nhu", "bin": "111010"},
    6: {"name": "Thiên Thủy Tụng", "bin": "010111"},
    7: {"name": "Địa Thủy Sư", "bin": "010000"},
    8: {"name": "Thủy Địa Tỷ", "bin": "000010"},
    9: {"name": "Phong Thiên Tiểu Súc", "bin": "111011"},
    10: {"name": "Thiên Trạch Lý", "bin": "110111"},
    11: {"name": "Địa Thiên Thái", "bin": "111000"},
    12: {"name": "Thiên Địa Bĩ", "bin": "000111"},
    13: {"name": "Thiên Hỏa Đồng Nhân", "bin": "101111"},
    14: {"name": "Hỏa Thiên Đại Hữu", "bin": "111101"},
    15: {"name": "Địa Sơn Khiêm", "bin": "001000"},
    16: {"name": "Lôi Địa Dự", "bin": "000100"},
    17: {"name": "Trạch Lôi Tùy", "bin": "100110"},
    18: {"name": "Sơn Phong Cổ", "bin": "011001"},
    19: {"name": "Địa Trạch Lâm", "bin": "110000"},
    20: {"name": "Phong Địa Quán", "bin": "000011"},
    21: {"name": "Hỏa Lôi Phệ Hạp", "bin": "100101"},
    22: {"name": "Sơn Hỏa Bí", "bin": "101001"},
    23: {"name": "Sơn Địa Bác", "bin": "000001"},
    24: {"name": "Địa Lôi Phục", "bin": "100000"},
    25: {"name": "Thiên Lôi Vô Vọng", "bin": "100111"},
    26: {"name": "Sơn Thiên Đại Súc", "bin": "111001"},
    27: {"name": "Sơn Lôi Di", "bin": "100001"},
    28: {"name": "Trạch Phong Đại Quá", "bin": "011110"},
    29: {"name": "Thuần Khảm", "bin": "010010"},
    30: {"name": "Thuần Ly", "bin": "101101"},
    31: {"name": "Trạch Sơn Hàm", "bin": "001110"},
    32: {"name": "Lôi Phong Hằng", "bin": "011100"},
    33: {"name": "Thiên Sơn Độn", "bin": "001111"},
    34: {"name": "Lôi Thiên Đại Tráng", "bin": "111100"},
    35: {"name": "Hỏa Địa Tấn", "bin": "000101"},
    36: {"name": "Địa Hỏa Minh Di", "bin": "101000"},
    37: {"name": "Phong Hỏa Gia Nhân", "bin": "101011"},
    38: {"name": "Hỏa Trạch Khuê", "bin": "110101"},
    39: {"name": "Thủy Sơn Kiển", "bin": "001010"},
    40: {"name": "Lôi Thủy Giải", "bin": "010100"},
    41: {"name": "Sơn Trạch Tổn", "bin": "110001"},
    42: {"name": "Phong Lôi Ích", "bin": "100011"},
    43: {"name": "Trạch Thiên Quải", "bin": "111110"},
    44: {"name": "Thiên Phong Cấu", "bin": "011111"},
    45: {"name": "Trạch Địa Tụy", "bin": "000110"},
    46: {"name": "Địa Phong Thăng", "bin": "011000"},
    47: {"name": "Trạch Thủy Khốn", "bin": "010110"},
    48: {"name": "Thủy Phong Tỉnh", "bin": "011010"},
    49: {"name": "Trạch Hỏa Cách", "bin": "101110"},
    50: {"name": "Hỏa Phong Đỉnh", "bin": "011101"},
    51: {"name": "Thuần Chấn", "bin": "100100"},
    52: {"name": "Thuần Cấn", "bin": "001001"},
    53: {"name": "Phong Sơn Tiệm", "bin": "001011"},
    54: {"name": "Lôi Trạch Quy Muội", "bin": "110100"},
    55: {"name": "Lôi Hỏa Phong", "bin": "101100"},
    56: {"name": "Hỏa Sơn Lữ", "bin": "001101"},
    57: {"name": "Thuần Tốn", "bin": "011011"},
    58: {"name": "Thuần Đoài", "bin": "110110"},
    59: {"name": "Phong Thủy Hoán", "bin": "010011"},
    60: {"name": "Thủy Trạch Tiết", "bin": "110010"},
    61: {"name": "Phong Trạch Trung Phu", "bin": "110011"},
    62: {"name": "Lôi Sơn Tiểu Quá", "bin": "001100"},
    63: {"name": "Thủy Hỏa Ký Tế", "bin": "101010"},
    64: {"name": "Hỏa Thủy Vị Tế", "bin": "010101"}
}

# Ánh xạ Bát Quái
BAGUA_MAP = {
    "111": ("Kiền", "Kim"), "000": ("Khôn", "Thổ"),
    "001": ("Cấn", "Thổ"), "110": ("Đoài", "Kim"),
    "010": ("Khảm", "Thủy"), "101": ("Ly", "Hỏa"),
    "100": ("Chấn", "Mộc"), "011": ("Tốn", "Mộc")
}

def get_can_chi(year):
    can_list = ["Canh", "Tân", "Nhâm", "Quý", "Giáp", "Ất", "Bính", "Đinh", "Mậu", "Kỷ"]
    chi_list = ["Thân", "Dậu", "Tuất", "Hợi", "Tý", "Sửu", "Dần", "Mão", "Thìn", "Tỵ", "Ngọ", "Mùi"]
    can = can_list[year % 10]
    chi = chi_list[year % 12]
    # Tý=1, Sửu=2 ...
    chi_num_map = {"Tý": 1, "Sửu": 2, "Dần": 3, "Mão": 4, "Thìn": 5, "Tỵ": 6, 
                   "Ngọ": 7, "Mùi": 8, "Thân": 9, "Dậu": 10, "Tuất": 11, "Hợi": 12}
    chi_num = chi_num_map[chi]
    # Âm Dương của năm dựa vào Can
    is_yang = year % 2 == 0 # Canh (0)->Dương, Tân (1)->Âm... 
    return can, chi, chi_num, is_yang

def calc_hoang_cuc(year):
    # Tính Tích Tuế (Mốc Giáp Tý 1984 = 10,155,901)
    tich_tue = 10155901 + (year - 1984)
    
    # Quẻ Trực Niên
    que_num = tich_tue % 64
    if que_num == 0: que_num = 64
    que = WEN_WANG_64[que_num]
    
    # Tính Hào Động
    can, chi, chi_num, is_yang = get_can_chi(year)
    bin_str = que["bin"] # "101011" (Hào 1 đến Hào 6)
    
    # Tính Hào Động theo Thái Ất Kể Năm
    target_lines = []
    if is_yang:
        # Năm Dương: đếm hào Dương từ dưới lên (H1 -> H6)
        for i in range(6):
            if bin_str[i] == '1': target_lines.append(i + 1)
    else:
        # Năm Âm: đếm hào Âm từ trên xuống (H6 -> H1)
        for i in range(5, -1, -1):
            if bin_str[i] == '0': target_lines.append(i + 1)
            
    # Fallback nếu quẻ không có Hào Dương/Âm tương ứng (Ví dụ Thuần Kiền năm Âm)
    if not target_lines:
        if is_yang: target_lines = [1, 2, 3, 4, 5, 6]
        else: target_lines = [6, 5, 4, 3, 2, 1]
                
    # Đếm đến Chi năm (1-12)
    idx = (chi_num - 1) % len(target_lines)
    hao_dong = target_lines[idx]
    
    # Thể Dụng (Quy ước: Hào động thuộc Thượng Quái (4,5,6) -> Thượng là Dụng, Hạ là Thể)
    # Hào động thuộc Hạ Quái (1,2,3) -> Hạ là Dụng, Thượng là Thể.
    ha_quai = bin_str[0:3]
    thuong_quai = bin_str[3:6]
    
    if hao_dong <= 3:
        the_quai = thuong_quai
        dung_quai = ha_quai
    else:
        the_quai = ha_quai
        dung_quai = thuong_quai
        
    the_name, the_hanh = BAGUA_MAP[the_quai]
    dung_name, dung_hanh = BAGUA_MAP[dung_quai]
    
    return {
        "year": year,
        "can_chi": f"{can} {chi} ({'Dương' if is_yang else 'Âm'})",
        "tich_tue": tich_tue,
        "que_num": que_num,
        "que_name": que["name"],
        "bin": bin_str,
        "chi_num": chi_num,
        "hao_dong": hao_dong,
        "the_hanh": the_hanh,
        "the_quai": the_name,
        "dung_hanh": dung_hanh,
        "dung_quai": dung_name
    }

if __name__ == "__main__":
    print("=== TEST HOÀNG CỰC KINH THẾ CÁC NĂM ===")
    for y in range(2021, 2027):
        res = calc_hoang_cuc(y)
        print(f"Năm {res['year']} ({res['can_chi']}): Quẻ {res['que_num']} {res['que_name']} | Động Hào {res['hao_dong']} | Thể Năm = {res['the_hanh']} ({res['the_quai']})")

