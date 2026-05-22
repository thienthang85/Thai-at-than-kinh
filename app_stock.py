import streamlit as st
import datetime
import pandas as pd
import re
from thai_at_calculator import ThaiAtCalculator
from borax.calendars.lunardate import LunarDate

st.set_page_config(page_title="Binh Pháp Đầu Tư Cổ Phiếu - Thái Ất", layout="wide")

# Khởi tạo Calculator
@st.cache_resource
def get_calculator():
    return ThaiAtCalculator()

calc = get_calculator()

# Helper Mappings
INDUSTRY_MAPPING = {
    "Ngân hàng": "Kim", "Thép": "Kim", "Bán lẻ": "Kim", "Chứng khoán": "Kim", "Cơ điện lạnh": "Kim",
    "Cảng biển": "Thủy", "Vận tải biển": "Thủy", "Đồ uống": "Thủy", "Thủy sản": "Thủy",
    "Bất động sản": "Thổ", "Bất động sản KCN": "Thổ", "Xây dựng": "Thổ", "Hạ tầng": "Thổ", "Thực phẩm": "Thổ", "Bánh kẹo": "Thổ",
    "Công nghệ": "Hỏa", "Hóa chất": "Hỏa", "Phân bón": "Hỏa", "Hàng không": "Hỏa", "Dược phẩm": "Hỏa",
    "Cao su": "Mộc", "Nông nghiệp": "Mộc", "Gỗ": "Mộc", "Nhựa": "Mộc"
}

TRIGRAM_ELEMENTS = {
    '111': ('Kiền', 'Kim'), '000': ('Khôn', 'Thổ'), '100': ('Chấn', 'Mộc'), '011': ('Tốn', 'Mộc'),
    '010': ('Khảm', 'Thủy'), '101': ('Ly', 'Hỏa'), '001': ('Cấn', 'Thổ'), '110': ('Đoài', 'Kim')
}

SINH = {'Mộc': 'Hỏa', 'Hỏa': 'Thổ', 'Thổ': 'Kim', 'Kim': 'Thủy', 'Thủy': 'Mộc'}
KHAC = {'Mộc': 'Thổ', 'Thổ': 'Thủy', 'Thủy': 'Hỏa', 'Hỏa': 'Kim', 'Kim': 'Mộc'}
HANH_MAP = {1: 'Kim', 2: 'Thủy', 3: 'Mộc', 4: 'Hỏa', 5: 'Thổ'}

# Bát Quái Tiên Thiên: 1=Kiền, 2=Đoài, 3=Ly, 4=Chấn, 5=Tốn, 6=Khảm, 7=Cấn, 8=Khôn
BAGUA_BIN = {1:'111', 2:'110', 3:'101', 4:'100', 5:'011', 6:'010', 7:'001', 8:'000'}
CAN_NUM = {'Giáp':1,'Ất':2,'Bính':3,'Đinh':4,'Mậu':5,'Kỷ':6,'Canh':7,'Tân':8,'Nhâm':9,'Quý':10}
CHI_NUM = {'Tý':1,'Sửu':2,'Dần':3,'Mão':4,'Thìn':5,'Tỵ':6,'Ngọ':7,'Mùi':8,'Thân':9,'Dậu':10,'Tuất':11,'Hợi':12}

def find_que_by_binary(binary_str):
    for k, v in calc.HEXAGRAMS.items():
        if v['binary'] == binary_str:
            return k
    return 0

# --- HOÀNG CỰC KINH THẾ (TÍCH TUẾ) ---
WEN_WANG_64 = {
    1: {"name": "Thuần Kiền", "bin": "111111"}, 2: {"name": "Thuần Khôn", "bin": "000000"},
    3: {"name": "Thủy Lôi Truân", "bin": "100010"}, 4: {"name": "Sơn Thủy Mông", "bin": "010001"},
    5: {"name": "Thủy Thiên Nhu", "bin": "111010"}, 6: {"name": "Thiên Thủy Tụng", "bin": "010111"},
    7: {"name": "Địa Thủy Sư", "bin": "010000"}, 8: {"name": "Thủy Địa Tỷ", "bin": "000010"},
    9: {"name": "Phong Thiên Tiểu Súc", "bin": "111011"}, 10: {"name": "Thiên Trạch Lý", "bin": "110111"},
    11: {"name": "Địa Thiên Thái", "bin": "111000"}, 12: {"name": "Thiên Địa Bĩ", "bin": "000111"},
    13: {"name": "Thiên Hỏa Đồng Nhân", "bin": "101111"}, 14: {"name": "Hỏa Thiên Đại Hữu", "bin": "111101"},
    15: {"name": "Địa Sơn Khiêm", "bin": "001000"}, 16: {"name": "Lôi Địa Dự", "bin": "000100"},
    17: {"name": "Trạch Lôi Tùy", "bin": "100110"}, 18: {"name": "Sơn Phong Cổ", "bin": "011001"},
    19: {"name": "Địa Trạch Lâm", "bin": "110000"}, 20: {"name": "Phong Địa Quán", "bin": "000011"},
    21: {"name": "Hỏa Lôi Phệ Hạp", "bin": "100101"}, 22: {"name": "Sơn Hỏa Bí", "bin": "101001"},
    23: {"name": "Sơn Địa Bác", "bin": "000001"}, 24: {"name": "Địa Lôi Phục", "bin": "100000"},
    25: {"name": "Thiên Lôi Vô Vọng", "bin": "100111"}, 26: {"name": "Sơn Thiên Đại Súc", "bin": "111001"},
    27: {"name": "Sơn Lôi Di", "bin": "100001"}, 28: {"name": "Trạch Phong Đại Quá", "bin": "011110"},
    29: {"name": "Thuần Khảm", "bin": "010010"}, 30: {"name": "Thuần Ly", "bin": "101101"},
    31: {"name": "Trạch Sơn Hàm", "bin": "001110"}, 32: {"name": "Lôi Phong Hằng", "bin": "011100"},
    33: {"name": "Thiên Sơn Độn", "bin": "001111"}, 34: {"name": "Lôi Thiên Đại Tráng", "bin": "111100"},
    35: {"name": "Hỏa Địa Tấn", "bin": "000101"}, 36: {"name": "Địa Hỏa Minh Di", "bin": "101000"},
    37: {"name": "Phong Hỏa Gia Nhân", "bin": "101011"}, 38: {"name": "Hỏa Trạch Khuê", "bin": "110101"},
    39: {"name": "Thủy Sơn Kiển", "bin": "001010"}, 40: {"name": "Lôi Thủy Giải", "bin": "010100"},
    41: {"name": "Sơn Trạch Tổn", "bin": "110001"}, 42: {"name": "Phong Lôi Ích", "bin": "100011"},
    43: {"name": "Trạch Thiên Quải", "bin": "111110"}, 44: {"name": "Thiên Phong Cấu", "bin": "011111"},
    45: {"name": "Trạch Địa Tụy", "bin": "000110"}, 46: {"name": "Địa Phong Thăng", "bin": "011000"},
    47: {"name": "Trạch Thủy Khốn", "bin": "010110"}, 48: {"name": "Thủy Phong Tỉnh", "bin": "011010"},
    49: {"name": "Trạch Hỏa Cách", "bin": "101110"}, 50: {"name": "Hỏa Phong Đỉnh", "bin": "011101"},
    51: {"name": "Thuần Chấn", "bin": "100100"}, 52: {"name": "Thuần Cấn", "bin": "001001"},
    53: {"name": "Phong Sơn Tiệm", "bin": "001011"}, 54: {"name": "Lôi Trạch Quy Muội", "bin": "110100"},
    55: {"name": "Lôi Hỏa Phong", "bin": "101100"}, 56: {"name": "Hỏa Sơn Lữ", "bin": "001101"},
    57: {"name": "Thuần Tốn", "bin": "011011"}, 58: {"name": "Thuần Đoài", "bin": "110110"},
    59: {"name": "Phong Thủy Hoán", "bin": "010011"}, 60: {"name": "Thủy Trạch Tiết", "bin": "110010"},
    61: {"name": "Phong Trạch Trung Phu", "bin": "110011"}, 62: {"name": "Lôi Sơn Tiểu Quá", "bin": "001100"},
    63: {"name": "Thủy Hỏa Ký Tế", "bin": "101010"}, 64: {"name": "Hỏa Thủy Vị Tế", "bin": "010101"}
}

BAGUA_MAP_HC = {
    "111": ("Kiền", "Kim"), "000": ("Khôn", "Thổ"),
    "001": ("Cấn", "Thổ"), "110": ("Đoài", "Kim"),
    "010": ("Khảm", "Thủy"), "101": ("Ly", "Hỏa"),
    "100": ("Chấn", "Mộc"), "011": ("Tốn", "Mộc")
}

def calc_macro_hoang_cuc(year):
    can_list = ["Canh", "Tân", "Nhâm", "Quý", "Giáp", "Ất", "Bính", "Đinh", "Mậu", "Kỷ"]
    chi_list = ["Thân", "Dậu", "Tuất", "Hợi", "Tý", "Sửu", "Dần", "Mão", "Thìn", "Tỵ", "Ngọ", "Mùi"]
    can = can_list[year % 10]
    chi = chi_list[year % 12]
    chi_num_map = {"Tý": 1, "Sửu": 2, "Dần": 3, "Mão": 4, "Thìn": 5, "Tỵ": 6, 
                   "Ngọ": 7, "Mùi": 8, "Thân": 9, "Dậu": 10, "Tuất": 11, "Hợi": 12}
    chi_num = chi_num_map[chi]
    is_yang = year % 2 == 0 
    
    tich_tue = 10155901 + (year - 1984)
    que_num = tich_tue % 64 or 64
    que = WEN_WANG_64[que_num]
    bin_str = que["bin"] 
    
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
    
    ha_quai = bin_str[0:3]
    thuong_quai = bin_str[3:6]
    
    if hao_dong <= 3:
        the_quai, dung_quai = thuong_quai, ha_quai
    else:
        the_quai, dung_quai = ha_quai, thuong_quai
        
    the_name, the_hanh = BAGUA_MAP_HC[the_quai]
    dung_name, dung_hanh = BAGUA_MAP_HC[dung_quai]
    
    return {
        'can': can, 'chi': chi, 'que_num': que_num, 'hd': hao_dong,
        'que_name': que["name"], 'the_nam': the_hanh, 'the_quai': the_name,
        'bin': bin_str, 'dung_hanh': dung_hanh, 'dung_quai': dung_name
    }

def get_relation(dn_hanh, vm_hanh):
    if dn_hanh == vm_hanh: return "VƯỢNG"
    if SINH.get(vm_hanh) == dn_hanh: return "TƯỚNG"
    if SINH.get(dn_hanh) == vm_hanh: return "THÔI"
    if KHAC.get(vm_hanh) == dn_hanh: return "TỬ"
    if KHAC.get(dn_hanh) == vm_hanh: return "TÙ"
    return "?"

def get_the_dung_score(the_bin, dung_bin):
    _, the_elem = TRIGRAM_ELEMENTS[the_bin]
    _, dung_elem = TRIGRAM_ELEMENTS[dung_bin]
    if the_elem == dung_elem: return 2, 'Tỷ hòa (Đồng hành)'
    elif SINH.get(dung_elem) == the_elem: return 5, 'Dụng sinh Thể (Đại Cát)'
    elif SINH.get(the_elem) == dung_elem: return -2, 'Thể sinh Dụng (Hao tổn)'
    elif KHAC.get(the_elem) == dung_elem: return 0, 'Thể khắc Dụng (Gian nan nhưng đoạt tài)'
    elif KHAC.get(dung_elem) == the_elem: return -5, 'Dụng khắc Thể (Đại Hung)'
    return 0, 'Bình Hòa'

def get_hex_the_dung(que_binary, hao_num):
    # Trigrams are lower (0-3) and upper (3-6)
    if hao_num in [1, 2, 3]:
        # Moving line is in lower trigram, so lower is Dụng, upper is Thể
        return que_binary[3:6], que_binary[0:3]
    else:
        # Moving line is in upper trigram, so upper is Dụng, lower is Thể
        return que_binary[0:3], que_binary[3:6]

def que_name(num):
    if num in calc.HEXAGRAMS:
        return f"Q{num}: {calc.HEXAGRAMS[num]['name']}"
    return f"Q{num}"

def que_quality(num):
    q = {1:"Cát",2:"Hung",3:"Hung",4:"Hung",5:"Trung",6:"Hung",7:"Trung",8:"Cát",
          9:"Cát",10:"Trung",11:"Đại Cát",12:"Đại Hung",13:"Cát",14:"Cát",15:"Cát",16:"Cát",
          17:"Trung",18:"Hung",19:"Đại Cát",20:"Trung",21:"Hung",22:"Trung",23:"Đại Hung",24:"Cát",
          25:"Hung",26:"Cát",27:"Trung",28:"Hung",29:"Đại Hung",30:"Cát",31:"Cát",32:"Cát",
          33:"Trung",34:"Trung",35:"Cát",36:"Hung",37:"Cát",38:"Trung",39:"Hung",40:"Cát",
          41:"Trung",42:"Cát",43:"Hung",44:"Trung",45:"Cát",46:"Cát",47:"Hung",48:"Cát",
          49:"Trung",50:"Trung",51:"Trung",52:"Trung",53:"Cát",54:"Hung",55:"Cát",56:"Hung",
          57:"Cát",58:"Cát",59:"Hung",60:"Trung",61:"Cát",62:"Trung",63:"Cát",64:"Hung"}
    return q.get(num, "?")

def get_cung_element(year, month, is_lunar=False):
    # Returns Cụm name, Element, and Season Year
    if is_lunar:
        if month in [2, 3, 4]:
            cum_key = "Cụm1_T234"
            season_year = year
        elif month in [5, 6, 7]:
            cum_key = "Cụm2_T567"
            season_year = year
        elif month in [8, 9, 10]:
            cum_key = "Cụm3_T8910"
            season_year = year
        elif month in [11, 12]:
            cum_key = "Cụm4_T11121"
            season_year = year
        else:  # month == 1
            cum_key = "Cụm4_T11121"
            season_year = year - 1
        cungs = calc.get_macro_monthly_cung(season_year)
        return cum_key, cungs[cum_key][2]
    else:
        # Tiết khí calendar mapping roughly
        if month in [2, 3, 4]:
            cum_key = "Cụm1_T234"
        elif month in [5, 6, 7]:
            cum_key = "Cụm2_T567"
        elif month in [8, 9, 10]:
            cum_key = "Cụm3_T8910"
        else:
            cum_key = "Cụm4_T11121"
            if month == 1: year = year - 1
        cungs = calc.get_macro_monthly_cung(year)
        return cum_key, cungs[cum_key][2]

# PRESET DATABASE
PRESET_STOCKS = {
    "ACB": {"date": datetime.date(2006, 11, 21), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "BBC": {"date": datetime.date(2001, 12, 19), "time": datetime.time(9, 0), "ind": "Bánh kẹo"},
    "BCM": {"date": datetime.date(2018, 3, 21), "time": datetime.time(9, 0), "ind": "Bất động sản"},
    "BID": {"date": datetime.date(2014, 1, 24), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "CII": {"date": datetime.date(2006, 5, 18), "time": datetime.time(9, 0), "ind": "Hạ tầng"},
    "CTG": {"date": datetime.date(2009, 7, 16), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "DCM": {"date": datetime.date(2015, 3, 31), "time": datetime.time(9, 0), "ind": "Phân bón"},
    "DGC": {"date": datetime.date(2014, 8, 13), "time": datetime.time(9, 0), "ind": "Hóa chất"},
    "DPG": {"date": datetime.date(2017, 1, 12), "time": datetime.time(9, 0), "ind": "Xây dựng"},
    "FPT": {"date": datetime.date(2006, 12, 13), "time": datetime.time(9, 0), "ind": "Công nghệ"},
    "GMD": {"date": datetime.date(2002, 4, 22), "time": datetime.time(9, 0), "ind": "Cảng biển"},
    "GVR": {"date": datetime.date(2018, 3, 21), "time": datetime.time(9, 0), "ind": "Cao su"},
    "HAH": {"date": datetime.date(2015, 3, 11), "time": datetime.time(9, 0), "ind": "Vận tải biển"},
    "HAR": {"date": datetime.date(2013, 1, 17), "time": datetime.time(9, 0), "ind": "Bất động sản"},
    "HDB": {"date": datetime.date(2018, 1, 5), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "HPG": {"date": datetime.date(2007, 11, 15), "time": datetime.time(9, 0), "ind": "Thép"},
    "IDC": {"date": datetime.date(2017, 11, 24), "time": datetime.time(9, 0), "ind": "Bất động sản KCN"},
    "KBC": {"date": datetime.date(2007, 12, 18), "time": datetime.time(9, 0), "ind": "Bất động sản KCN"},
    "KDC": {"date": datetime.date(2005, 12, 12), "time": datetime.time(9, 0), "ind": "Thực phẩm"},
    "LPB": {"date": datetime.date(2017, 10, 24), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "MBB": {"date": datetime.date(2011, 11, 1), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "MSN": {"date": datetime.date(2009, 11, 5), "time": datetime.time(9, 0), "ind": "Thực phẩm"},
    "MWG": {"date": datetime.date(2014, 7, 14), "time": datetime.time(9, 0), "ind": "Bán lẻ"},
    "ORS": {"date": datetime.date(2010, 7, 12), "time": datetime.time(9, 0), "ind": "Chứng khoán"},
    "REE": {"date": datetime.date(2000, 7, 28), "time": datetime.time(9, 0), "ind": "Cơ điện lạnh"},
    "SAB": {"date": datetime.date(2016, 12, 6), "time": datetime.time(9, 0), "ind": "Đồ uống"},
    "SGP": {"date": datetime.date(2016, 4, 25), "time": datetime.time(9, 0), "ind": "Cảng biển"},
    "SHB": {"date": datetime.date(2009, 4, 20), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "SIP": {"date": datetime.date(2019, 6, 6), "time": datetime.time(9, 0), "ind": "Bất động sản KCN"},
    "SSI": {"date": datetime.date(2006, 12, 15), "time": datetime.time(9, 0), "ind": "Chứng khoán"},
    "STB": {"date": datetime.date(2006, 7, 12), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "TCH": {"date": datetime.date(2016, 10, 5), "time": datetime.time(9, 0), "ind": "Bất động sản"},
    "VCB": {"date": datetime.date(2009, 6, 30), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "VHM": {"date": datetime.date(2018, 5, 17), "time": datetime.time(9, 0), "ind": "Bất động sản"},
    "VIB": {"date": datetime.date(2017, 1, 9), "time": datetime.time(9, 0), "ind": "Ngân hàng"},
    "VIC": {"date": datetime.date(2007, 9, 19), "time": datetime.time(9, 0), "ind": "Bất động sản"},
    "VJC": {"date": datetime.date(2017, 2, 28), "time": datetime.time(9, 0), "ind": "Hàng không"},
    "VNM": {"date": datetime.date(2006, 1, 19), "time": datetime.time(9, 0), "ind": "Thực phẩm"},
    "VRE": {"date": datetime.date(2017, 11, 6), "time": datetime.time(9, 0), "ind": "Bất động sản"}
}

# UI 
st.title("🎯 TRẬN ĐỒ MA TRẬN CK")

st.sidebar.header("📥 NHẬP LIỆU CỔ PHIẾU")

ticker_options = ["Tùy chỉnh (Mã khác)"] + list(PRESET_STOCKS.keys())
selected_preset = st.sidebar.selectbox("Chọn Mã Cổ Phiếu (Top 39)", ticker_options)

industry_options = list(INDUSTRY_MAPPING.keys()) + ["Khác"]

if selected_preset != "Tùy chỉnh (Mã khác)":
    ticker = selected_preset
    preset_data = PRESET_STOCKS[ticker]
    st.sidebar.success(f"✅ Đã tự động điền dữ liệu cho {ticker}")
    industry = st.sidebar.selectbox("Nhóm Ngành", industry_options, index=industry_options.index(preset_data["ind"]), disabled=True)
    first_trade_date = st.sidebar.date_input("Ngày Giao Dịch Đầu Tiên", preset_data["date"], disabled=True)
    first_trade_time = st.sidebar.time_input("Giờ Lên Sàn", preset_data["time"], disabled=True)
else:
    ticker = st.sidebar.text_input("Nhập Mã Cổ Phiếu", "VHM").upper()
    industry = st.sidebar.selectbox("Nhóm Ngành", industry_options, index=industry_options.index("Bất động sản"))
    first_trade_date = st.sidebar.date_input("Ngày Giao Dịch Đầu Tiên", datetime.date(2018, 5, 17))
    first_trade_time = st.sidebar.time_input("Giờ Lên Sàn", datetime.time(9, 0))

if industry == "Khác":
    custom_hanh = st.sidebar.selectbox("Tự chọn Hành Ngành", ["Kim", "Thủy", "Mộc", "Hỏa", "Thổ"])
    ind_hanh = custom_hanh
else:
    ind_hanh = INDUSTRY_MAPPING[industry]

macro_year = st.sidebar.number_input("Năm Vĩ Mô", min_value=2000, max_value=2050, value=2026, step=1)
calendar_type = st.sidebar.radio("Loại Lịch Tính", ["Tiết Khí (Khuyên dùng)", "Âm Lịch"])
is_lunar = (calendar_type == "Âm Lịch")

# TỰ ĐỘNG TÍNH TOÁN CHẾ ĐỘ THỊ TRƯỜNG: HOÀNG CỰC KINH THẾ
macro_vm = calc_macro_hoang_cuc(macro_year)
vm_que_num = macro_vm['que_num']
vm_que_name = macro_vm['que_name']
vm_qual = que_quality(vm_que_num)
the_nam = macro_vm['the_nam']
vm_bien = 0

# Đánh giá Market Mode bằng cả HÌNH TƯỢNG và THỂ DỤNG Quẻ Vĩ Mô
macro_bin = macro_vm['bin']
macro_hd = macro_vm['hd']
the_bin_macro = macro_bin[3:6] if macro_hd <= 3 else macro_bin[0:3]
dung_bin_macro = macro_bin[0:3] if macro_hd <= 3 else macro_bin[3:6]
macro_score, macro_td_label = get_the_dung_score(the_bin_macro, dung_bin_macro)

if vm_qual in ["Đại Cát", "Cát"]:
    if macro_score == 5:
        market_mode = "Down-trend (Bão Hòa Kỳ Vọng - Phân Phối Đỉnh) 📉"
    elif macro_score == -2:
        market_mode = "Sideway-trend (Đi Lên Trong Nghi Ngờ - Bị Rút Máu) ⚖️"
    elif macro_score == -5:
        market_mode = "Down-trend (Vỡ Trận Bất Ngờ) 📉"
    else: # 2, 0
        market_mode = "Up-trend (Dòng Tiền Đồng Thuận - Tăng Lành Mạnh) 📈"
elif vm_qual == "Trung":
    if macro_score in [5, 2, 0]:
        market_mode = "Up-trend (Đẩy Giá Khéo Léo - Đi Lên Vững Chắc) 📈"
    elif macro_score == -2:
        market_mode = "Sideway-trend (Giằng Co - Hao Tổn Nội Lực) ⚖️"
    else:
        market_mode = "Down-trend (Thiếu Thanh Khoản - Rút Tiền) 📉"
else: # Hung, Đại Hung
    if macro_score == -5:
        market_mode = "Down-trend (Suy Thoái - Khủng Hoảng Thực Sự) 📉"
    elif macro_score == -2:
        market_mode = "Sideway-trend (Thận Trọng Tích Lũy Bất Chấp Tin Xấu) ⚖️"
    else: # 5, 2, 0
        market_mode = "Up-trend (Bẫy Rũ Bỏ - Tạo Đáy Bằng Sự Hoảng Loạn) 📈"


st.sidebar.markdown("---")
if st.sidebar.button("Phân Tích Cổ Phiếu", type="primary", use_container_width=True):
    # Tính Bát tự & Bản mệnh
    dt_trade = datetime.datetime.combine(first_trade_date, first_trade_time)
    bt = calc.calculate_bat_tu(dt_trade)
    hanh_idx = calc.calculate_nap_am_ngu_hanh(bt['can_nam'], bt['chi_nam'])
    hanh_nap_am = HANH_MAP.get(hanh_idx, "?")
    
    # Tính Quẻ Dựng Nghiệp
    que_dn_info = calc.calculate_que_dich(bt)
    match = re.search(r'Quẻ (\d+)', que_dn_info['que_goc'])
    que_dn_num = int(match.group(1)) if match else 0
    
    # Tính Quẻ Lưu Niên
    tuoi = macro_year - first_trade_date.year + 1
    que_ln_num = (que_dn_num + tuoi) % 64 or 64
    hd_num, _, _ = calc.calculate_hao_dong_luu_nien(que_ln_num, macro_year)
    
    # Chấm điểm Thể Dụng Quẻ Lưu Niên
    que_binary = calc.HEXAGRAMS[que_ln_num]['binary']
    the_bin, dung_bin = get_hex_the_dung(que_binary, hd_num)
    score_orig, label_orig = get_the_dung_score(the_bin, dung_bin)
    
    # Tính Quẻ Biến Lưu Niên
    b_list = list(que_binary)
    b_list[hd_num - 1] = '1' if b_list[hd_num - 1] == '0' else '0'
    new_binary = ''.join(b_list)
    bien_num = 0
    for k, v in calc.HEXAGRAMS.items():
        if v['binary'] == new_binary:
            bien_num = k
            break
            
    the_bin_bien, dung_bin_bien = get_hex_the_dung(new_binary, hd_num)
    score_bien, label_bien = get_the_dung_score(the_bin_bien, dung_bin_bien)
    
    total_score = round(0.4 * score_orig + 0.6 * score_bien, 2)
    
    if total_score >= 3.5:
        rank_label = "Đại Cát 🌟🌟🌟"
    elif total_score >= 1.2:
        rank_label = "Cát 🌟🌟"
    elif total_score >= 0.2:
        rank_label = "Tiểu Cát 🌟"
    elif total_score > -0.2:
        rank_label = "Bình Hòa ⚖️"
    elif total_score > -1.2:
        rank_label = "Tiểu Hung ⚠️"
    elif total_score > -3.5:
        rank_label = "Hung ❌"
    else:
        rank_label = "Đại Hung 💀"
    
    # Hiển thị Phần 1: Hồ sơ nội tại
    st.subheader(f"1. HỒ SƠ NỘI TẠI DOANH NGHIỆP: {ticker}")
    c1, c2, c3, c4 = st.columns(4)
    c1.info(f"**Bản Mệnh Cổ Phiếu:**\n\nNạp âm: **{hanh_nap_am}**")
    c2.info(f"**Hành Của Ngành:**\n\n**{ind_hanh}** ({industry})")
    c3.success(f"**Quẻ Dựng Nghiệp:**\n\n{que_name(que_dn_num)}")
    c4.success(f"**Quẻ Lưu Niên {macro_year}:**\n\n{que_name(que_ln_num)} (Động Hào {hd_num})")
    
    st.markdown(f"> **Đánh giá Thể Dụng Quẻ Năm:** Tổng điểm **{total_score}** ({rank_label})")
    st.markdown(f"> *Quẻ chính ({score_orig}đ): {label_orig} | Quẻ biến ({score_bien}đ): {label_bien} | Trọng số: 40% Chính - 60% Biến*")
    
    if total_score < -1.2:
        st.error("⚠️ Quẻ Lưu Niên yếu/xấu (Hung). Nguy cơ doanh nghiệp thua lỗ, sụp đổ từ nội tại (không có cơ bản). Thận trọng khi lướt sóng!")
    elif total_score >= 3.5:
        st.warning("🔥 Quẻ Lưu Niên Đại Cát. Doanh nghiệp ở đỉnh cao bùng nổ, biên độ lợi nhuận cực lớn. Tuy nhiên cần cẩn trọng nếu Vĩ mô đang ở cuối chu kỳ (phân phối đỉnh)!")
    elif total_score >= 0.2:
        st.success("✅ Quẻ Lưu Niên ổn định (Cát/Tiểu Cát). Doanh nghiệp có nội lực tốt, an toàn giải ngân!")
    else:
        st.info("⚖️ Quẻ Lưu Niên Bình Hòa. Doanh nghiệp đi ngang tích lũy, phù hợp mua gom chờ thời.")
    
    st.markdown("---")
    st.subheader(f"2. BẢNG CHIẾN LƯỢC 12 THÁNG NĂM {macro_year}")
    
    st.info(f"📈 **Hoàng Cực Kinh Thế ({macro_vm['can']} {macro_vm['chi']}):** Trực niên **{que_name(vm_que_num)}** Động H{macro_vm['hd']} | Thể Năm: **{the_nam} ({macro_vm['the_quai']})** ➔ **{market_mode}**")
    
    # KÍCH HOẠT LAYER 1: BỘ LỌC ĐẠI CỤC NĂM (Ngành vs Thể Năm)
    year_sk_industry = get_relation(ind_hanh, the_nam)
    
    st.markdown(f"**LỚP 1 - TẦM NHÌN NĂM:** Ngành **{ind_hanh}** so với Vĩ mô Năm **{the_nam}** ➔ Trạng thái: **{year_sk_industry}**")
    if year_sk_industry == "TỬ":
        st.markdown("> *Ngành bị Vĩ mô khắc. Dù có sóng ngắn hạn cũng không bền, xu hướng chính là bị rút tiền.*")
    elif year_sk_industry == "TÙ":
        st.markdown("> *Ngành đi khắc Vĩ mô (Hút máu). Trong Down-trend biến thành Hầm trú ẩn hút tiền phòng thủ. Trong Up-trend dễ thành Dẫn sóng hoặc Điểm nghẽn bong bóng.*")
    elif year_sk_industry == "TƯỚNG":
        st.markdown("> *Vĩ mô sinh Ngành. Hưởng lợi thiên thời vĩ đại, đón dòng tiền khổng lồ đổ vào.*")
    elif year_sk_industry == "VƯỢNG":
        st.markdown("> *Hòa hợp Vĩ mô. Phát triển ổn định và cộng sinh theo chu kỳ chung.*")
    elif year_sk_industry == "THÔI":
        st.markdown("> *Ngành phải sinh ra Vĩ mô. Bị vắt kiệt sức lực để hỗ trợ thị trường, đà tăng trưởng suy yếu dần.*")
        
    st.markdown("**LỚP 2 - TẦM NHÌN LƯU NGUYỆT 3 THÁNG:** *(Chi tiết bảng chiến lược bên dưới - So găng Ngành/Doanh nghiệp với Hành Cụm Tháng)*")
    
    with st.expander("📚 BÍ KÍP ĐỌC TRẬN ĐỒ: MÔ HÌNH TÀI CHÍNH HÀNH VI ĐỊNH LƯỢNG (CLICK ĐỂ XEM)"):
        if "Down-trend" in market_mode:
            st.markdown("""
            **🔍 MÔI TRƯỜNG VĨ MÔ YẾU (Thiếu Thanh Khoản):**
            *   **🛑 CẮT MÁU / THÁO CHẠY:** Quẻ Tháng **HUNG**. Thị trường yếu + Tin xấu = Gãy thật, thủng đáy. Tuyệt đối không bắt đáy.
            *   **✅ MUA GOM DEEP VALUE:** Quẻ Tháng **CÁT** + Nạp âm **TỬ / THÔI**. Nội lực tốt nhưng giá bị ép xuống cực đại (kiệt cung). Chờ bật chữ V.
            *   **⚠️ PHÂN PHỐI ĐỈNH:** Nạp âm **VƯỢNG**. Mọi kỳ vọng tốt nhất đã ra (Priced-in), hết động lực tăng giá.
            *   **🛡️ TRÚ ẨN BẢO TOÀN:** Ngành **TƯỚNG**. Ngành được Vĩ mô bơm máu, phòng thủ an toàn trong giông bão.
            """)
        else:
            st.markdown("""
            **🚀 MÔI TRƯỜNG VĨ MÔ MẠNH (Thừa Thanh Khoản):**
            *   **🌪️ BẪY RŨ BỎ / MÚC BUY-THE-DIP:** Quẻ Tháng **HUNG**. Sập chớp nhoáng (Liquidity Shock) rũ sạch Margin rồi kéo trần V-shape. Cơ hội bắt đáy sinh lời lớn nhất.
            *   **🔥 SIÊU ĐẦU CƠ TỐI ĐA ALPHA:** Quẻ Tháng **CÁT** + Nạp âm **TỬ / THÔI**. Lõi tốt nén chặt, bung nổ nhờ tiền rẻ.
            *   **🚀 NGÀNH HÚT MÁU / DẪN SÓNG:** Ngành **TÙ**. Ngành lấy nguồn lực từ vĩ mô để mở rộng bong bóng.
            *   **⚠️ BÃO HÒA KỲ VỌNG:** Nạp âm **VƯỢNG** hoặc Ngành **VƯỢNG**. Cạn tiền đẩy mới, chỉ canh chốt lời.
            *   **🐢 CHẬM CHẠP / LỠ SÓNG:** Ngành **TƯỚNG**. Tiền đầu cơ không vào ngành phòng thủ, lỡ nhịp thị trường.
            """)
    
    co_so = que_ln_num + 2
    rows = []
    
    for month in range(1, 13):
        cum_key, vm_hanh = get_cung_element(macro_year, month, is_lunar=is_lunar)
        
        sk_industry = get_relation(ind_hanh, vm_hanh)
        sk_nap_am = get_relation(hanh_nap_am, vm_hanh)
        
        qt = (co_so + month) % 64 or 64
        q_qual = que_quality(qt)
        
        # KÍCH HOẠT LAYER 2: LOGIC KHUYẾN NGHỊ TÍCH HỢP DUAL-FILTER
        recommendation = "➖ Quan sát"
        if "Down-trend" in market_mode:
            if q_qual in ["Đại Hung", "Hung"]:
                recommendation = "🛑 CẮT MÁU / THÁO CHẠY (Thủng đáy)"
            elif sk_nap_am == "TỬ":
                recommendation = "🛑 TRÁNH XA (Dao rơi)"
            elif sk_nap_am == "VƯỢNG":
                recommendation = "🟡 PHÂN PHỐI ĐỈNH (Hết vị)"
            elif q_qual in ["Đại Cát", "Cát"] and sk_nap_am in ["TỬ", "THÔI"]:
                recommendation = "✅ MUA GOM DEEP VALUE (Chờ V)"
            else:
                recommendation = "⚠️ BULL-TRAP / HỒI KỸ THUẬT"
        else:
            if q_qual in ["Đại Hung", "Hung"]:
                recommendation = "🌪️ BẪY RŨ BỎ / BUY-THE-DIP (Chờ Volume bùng nổ)"
            elif sk_industry == "TÙ":
                recommendation = "🚀 DẪN SÓNG / SIÊU ĐẦU CƠ (All-in)"
            elif q_qual in ["Đại Cát", "Cát"] and sk_industry == "VƯỢNG":
                recommendation = "🟡 BÃO HÒA KỲ VỌNG / CHỐT LỜI"
            elif sk_nap_am == "TỬ":
                recommendation = "🛑 TRÁNH XA (Lõi cực yếu)"
            else:
                recommendation = "✅ TÍCH LŨY / HOLD"
                
        rows.append({
            "Tháng (AL)" if is_lunar else "Tháng (DL)": month,
            "Hành Vĩ Mô (Cụm)": f"{vm_hanh} ({cum_key[:4]})",
            "Thiên Thời (Ngành)": sk_industry,
            "Nhân Hòa 1 (Nạp Âm)": sk_nap_am,
            "Quẻ Tháng (Nhân Hòa 2)": f"{que_name(qt)} ({q_qual})",
            "Khuyến Nghị (Auto)": recommendation
        })
        
    df = pd.DataFrame(rows)
    
    # Apply style to table for highlighting
    def color_recommendation(val):
        # Mặc định (nếu không khớp)
        color = '#ffffff' 
        
        # NHÓM 1: MUA (Màu Xanh Lá)
        if any(kw in val for kw in ["🚀", "🔥", "🌪️", "✅", "DẪN SÓNG", "ĐẦU CƠ", "BUY-THE-DIP", "MUA GOM"]):
            color = '#00ff00' # Xanh lá (Mua)
            
        # NHÓM 2: BÁN (Màu Vàng)
        elif any(kw in val for kw in ["⚠️", "🛑", "PHÂN PHỐI", "BÃO HÒA", "CẮT MÁU", "BULL-TRAP"]):
            color = '#ffd700' # Vàng (Bán)
            
        # NHÓM 3: TRÁNH XA (Màu Đỏ)
        elif any(kw in val for kw in ["➖", "🐢", "🛡️", "Quan sát", "CHẬM CHẠP", "TRÚ ẨN"]):
            color = '#ff4b4b' # Đỏ (Tránh xa)
            
        return f'color: {color}; font-weight: bold'
        
    st.dataframe(df.style.map(color_recommendation, subset=['Khuyến Nghị (Auto)']), use_container_width=True, height=450)
    
    st.markdown(f"> **Giải thích chiến thuật ({market_mode}):**")
    if "Down-trend" in market_mode:
        st.markdown("*Đội lái đánh nghịch hành: Các tháng Vượng là phân phối đỉnh. Tập trung gom hàng khi cổ phiếu chịu áp lực TỬ (khắc nhập) hoặc THÔI (sinh xuất) vào những tháng có Quẻ Cát/Trung mà giá bị đè. Bán thẳng tay khi ra Quẻ Hung mà giá bị kéo thốc (Bẫy bull-trap).*")
    else:
        st.markdown("*Thị trường thuận sóng lớn: Tập trung các ngành TÙ (chiếm lĩnh vĩ mô) hoặc THÔI. Mua khi có Quẻ Cát. Bán khi Quẻ Hung xuất hiện vì dòng tiền tự nhiên sẽ tháo chạy.*")
    
    st.markdown("---")
    st.caption("Công cụ tính toán tự động dựa trên phân tích lượng tử kết hợp Thái Ất Dịch Lý & Tài chính hành vi (By Antigravity).")
