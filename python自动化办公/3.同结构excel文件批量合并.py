import os
import pandas as pd


def getAllFilesList(filepath):
    """
    获取指定目录下的所有xlsx文件列表
    :param filepath: 指定目录
    :return: 指定目录下的所有xlsx文件列表
    """
    files = []
    for file in os.listdir(filepath):
        if file.endswith(".xlsx"):
            files.append(filepath + file)

    return files


if __name__ == '__main__':
    filepath = "D:" + os.sep + "sampple" + os.sep + "xlsx_merge" + os.sep
    files = getAllFilesList(filepath)

    # 定义一个空的dataframe
    data = pd.DataFrame()

    for file in files:
        df = pd.read_excel(file)
        df_len = len(df)
        data = data.append(df)
        print('读取%i行数据，合并后文件%i列，名称：%s' % (df_len, len(data.columns), file.split('/')[-1]))

    # 重置索引
    data.reset_index(drop=True, inplace=True)
    # 查看数据
    print(data)
