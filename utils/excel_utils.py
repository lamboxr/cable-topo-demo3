def col_to_num(col):
    """将字母列（如"A"、"AB"）转换为数字（如1、28）"""
    num = 0
    for c in col:
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num


def num_to_col(num):
    """将数字（如1、28）转换为字母列（如"A"、"AB"）"""
    col = ""
    while num > 0:
        num -= 1  # 调整为0开始的索引
        col = chr(num % 26 + ord('A')) + col
        num = num // 26
    return col


def get_left_col_letter(current_col: str, i=1):
    """获取字母列左侧的列"""
    current_num = col_to_num(current_col)
    if current_num - 1 <= 0:
        return None
    return num_to_col(current_num - i)


def get_right_col_letter(current_col: str, i=1):
    """获取字母列左侧的列"""
    current_num = col_to_num(current_col)
    if current_num + 1 <= 0:
        return None
    return num_to_col(current_num + i)


if __name__ == '__main__':
    # 测试
    print(get_left_col_letter("C"))  # 输出："B"（C左侧是B）
    print(get_left_col_letter("AA"))  # 输出："Z"（AA左侧是Z）
    print(get_left_col_letter("AB"))  # 输出："AA"（AB左侧是AA）
    print(get_left_col_letter("A"))  # 输出：None（A无左侧）
