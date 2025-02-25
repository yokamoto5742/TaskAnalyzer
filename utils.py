from datetime import datetime


def parse_date_safely(date_str, format_str='%Y-%m-%d'):
    """
    文字列を安全にdatetimeオブジェクトに変換する
    
    Parameters:
    -----------
    date_str : str
        変換する日付文字列
    format_str : str
        日付のフォーマット
        
    Returns:
    --------
    datetime or None
        変換に成功した場合はdatetimeオブジェクト、失敗した場合はNone
    """
    try:
        return datetime.strptime(date_str, format_str)
    except (ValueError, TypeError):
        return None


def safe_float_conversion(value):
    """
    値を安全にfloatに変換する
    
    Parameters:
    -----------
    value : any
        変換する値
        
    Returns:
    --------
    float or None
        変換に成功した場合はfloat値、失敗した場合はNone
    """
    if isinstance(value, (int, float)):
        return float(value)
    elif isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            return None
    else:
        return None


def extract_name_from_content(content):
    """
    内容から名前を抽出する（括弧内の文字列）
    
    Parameters:
    -----------
    content : str
        抽出元の文字列
        
    Returns:
    --------
    tuple
        (name, cleaned_content)
    """
    import re
    name = None
    cleaned_content = content
    
    # 括弧内の文字列を抽出
    name_match = re.search(r'\((.*?)\)', content)
    if name_match:
        name = name_match.group(1)
        cleaned_content = re.sub(r'\(.*?\)', '', content).strip()
        
        # 最初の単語だけ取得
        if cleaned_content:
            cleaned_content = cleaned_content.split()[0]
            
    return name, cleaned_content
