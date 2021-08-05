import re


class CellStr2Number(object):

    def __init__(self, str_):
        if str_ is None:
            return
        if not isinstance(str_, str):
            return
        str_ = str_.strip()
        self.str_ = str_


    @staticmethod
    def str2Number(detect_str):
        try:
            return float(detect_str)
        except:
            return detect_str

    @staticmethod
    def floatPointNum(detect_float):
        if not isinstance(detect_float, (float, int)):
            raise TypeError(f'{detect_float}不是数字。')
        if isinstance(detect_float, int):
            return 0
        return len(str(detect_float).split('.')[1])

    def plusOrMinus(self):
        calcSign = ['+', '-']
        if not any(calc in self.str_ for calc in calcSign):
            CellStr2Number.str2Number(self.str_)
        try:
            splitSign = '+'
            strCopy_ = self.str_
            strCopy_ = strCopy_.replace('-', '+-')
            numberList = strCopy_.split(splitSign)
            calcNum = [CellStr2Number.str2Number(numStr_) for numStr_ in numberList]
            pointNum = max([CellStr2Number.floatPointNum(num) for num in calcNum])
            return round(sum(calcNum), pointNum)
        except:
            return self.str_



if __name__ == '__main__':
    print(CellStr2Number('15.2a-452.3').plusOrMinus())