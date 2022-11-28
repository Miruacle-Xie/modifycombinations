import os
import itertools

from openpyxl import load_workbook
from openpyxl import Workbook

import pandas as pd


def combi(seq):
    if not seq:
        yield []
    else:
        for element in seq[0]:
            for rest in combi(seq[1:]):
                yield [element] + rest



def main():
    filename = input("\n文件路径：\n")
    filename = filename.replace("\"", "").replace("\'", "")
    # print(os.path.splitext(os.path.realpath(filename)))
    df = pd.read_excel(filename, header=None)
    modifylist = []
    topiclist = []
    tmplist = []
    for i in range(len(df)):
        # print(df.iloc[i].isna().sum())
        # print(df.shape[1]-df.iloc[i].isna().sum())
        for j in range(1, df.shape[1]-df.iloc[i].isna().sum()):
            tmplist.append(df.iloc[i, j].split(","))
        productlist = list(itertools.product(*tmplist))
        productlist = [" ".join(i) for i in productlist]
        # print(productlist)
        modifylist.extend(productlist)
        # c = [df.iloc[i, 0] for cnt in range(len(productlist))]
        topiclist.extend([df.iloc[i, 0] for cnt in range(len(productlist))])
        tmplist = []
    df_result = pd.concat([pd.Series(topiclist), pd.Series(modifylist)], axis=1)
    print(df_result)
    print(df_result.shape)
    df_result.to_excel(os.path.splitext(os.path.realpath(filename))[0] + "-result" + ".xlsx", index=False, header=False)
    input("已生成报告, 按回车键结束")


if __name__ == "__main__":
    # print("1")
    main()
