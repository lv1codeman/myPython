import re


def is_contain_chinese(word):
    '''
    判断字符串是否包含中文字符
    :param word: 字符串
    :return: 布尔值，True表示包含中文，False表示不包含中文
    '''
    pattern = re.compile(r'[\u4e00-\u9fa5]')
    match = pattern.search(word)
    return True if match else False


def get_num_column_dict():
    '''
    產生dict  
    {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z', 27: 'AA', 28: 'AB', 29: 'AC', 30: 'AD', 31: 'AE', 32: 'AF', 33: 'AG', 34: 'AH', 35: 'AI', 36: 'AJ', 37: 'AK', 38: 'AL', 39: 'AM', 40: 'AN', 41: 'AO', 42: 'AP', 43: 'AQ', 44: 'AR', 45: 'AS', 46: 'AT', 47: 'AU', 48: 'AV', 49: 'AW', 50: 'AX', 51: 'AY', 52: 'AZ'}
    '''
    num_str_dict = {}
    A_Z = [chr(a) for a in range(ord('A'), ord('Z')+1)]
    AA_AZ = ['A' + chr(a) for a in range(ord('A'), ord('Z')+1)]
    A_AZ = A_Z + AA_AZ
    for i in A_AZ:
        num_str_dict[A_AZ.index(i) + 1] = i
    return num_str_dict


def getSyllabusColumns(result, hasCHT, hasENG):
    output = [''] * 3
    if hasCHT == False and hasENG == False:
        output[0] = "無檔案"
        output[1] = "No file"
        output[2] = "Y"
    else:
        output[2] = "N"
        if hasCHT == True and hasENG == False:
            output[0] = result.find("a").text.strip()
            output[1] = "No file"
        elif hasCHT == False and hasENG == True:
            output[0] = "無檔案"
            output[1] = result.find("a").text.strip()
        else:
            output[0] = result.find("a").text.strip()
            output[1] = result.find(
                "a").next_sibling.next_sibling.next_sibling.strip()
    return output

# print(get_num_column_dict.__doc__)
