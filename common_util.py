def covert_10_to_26(num):
    """
    进制转换 十进制转26进制
    参考：
    https://leetcode-cn.com/problems/excel-sheet-column-title/solution/excelbiao-lie-ming-cheng-by-leetcode-sol-hgj4/
    :param num:
    :return:
    """
    ans = list()
    while num > 0:
        num -= 1
        ans.append(chr(num % 26 + ord("A")))
        num //= 26
    return "".join(ans[::-1])


def covert_26_to_10(s: str):
    """
    进制转换 26进制转十进制
    参考：
    https://leetcode-cn.com/problems/excel-sheet-column-number/solution/hua-jie-suan-fa-171-excelbiao-lie-xu-hao-by-guanpe/
    :param s:
    :return:
    """
    ans = 0
    for i in range(len(s)):
        num = ord(s[i]) - ord('A') + 1
        ans = ans * 26 + num
    return ans
