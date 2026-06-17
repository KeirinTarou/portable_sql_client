def truncate(text: str, length: int = 200) -> str:
    """ 文字列を指定字数で切り捨てて`...`に置き換える
    
    :param text: 処理対象文字列
    :type text: str
    :param length: 最大文字数
    :type length: int
    :return: 変換後の文字列
    :rtype: str

    .. note::
    - app/core/text_utils.py
    """
    if len(text) <= length:
        return text
    return text[:length - 3] + "..."