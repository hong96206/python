import os
from win32com import client


def getAllFilesList(filepath):
    """
    获取指定目录下的所有xls文件列表
    :param filepath: 指定目录
    :return: 指定目录下的所有xls文件列表
    """
    files = []
    for file in os.listdir(filepath):
        if file.endswith(".xls"):
            files.append(filepath + file)

    return files


if __name__ == '__main__':
    filepath = "D:" + os.sep + "sampple" + os.sep + "xls" + os.sep
    # print(filepath)
    files = getAllFilesList(filepath)

    # 运行excel程序
    excel = client.gencache.EnsureDispatch('Excel.Application')
    # 设置excel另存时当有重名的文件时不提示弹窗预警
    excel.Application.DisplayAlerts = False


    i = 0
    for file in files:
        try:
            xls = excel.Workbooks.Open(file)  # 打开excel文件
            xls.SaveAs("{}x".format(file), 51)  # 另存为后缀为“.xlsx”的文件，其中参数51指的doc文件
            xls.Close()  # 关闭原来的excel文件
            print(file + '转换成功！')
            i += 1
        except:
            print(file + '转换失败！')
            files.append(file)  # 将文件名添加到files列表中重新读取
            pass

    print("转换文件%i个" % i)
    excel.Quit()
