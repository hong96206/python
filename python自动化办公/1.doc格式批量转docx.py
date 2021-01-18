import os
from win32com import client


def getAllFilesList(filepath):
    """
    获取指定目录下的所有doc文件列表
    :param filepath: 指定目录
    :return: 指定目录下的所有doc文件列表
    """
    files = []
    for file in os.listdir(filepath):
        if file.endswith(".doc"):
            files.append(filepath + file)

    return files


if __name__ == '__main__':
    filepath = "D:" + os.sep + "sampple" + os.sep + "doc" + os.sep
    # print(filepath)
    files = getAllFilesList(filepath)

    word = client.Dispatch("Word.Application")

    i = 0
    for file in files:
        try:
            doc = word.Documents.Open(file)  # 打开word文件
            doc.SaveAs("{}x".format(file), 12)  # 另存为后缀为“.docx”的文件，其中参数12指的doc文件
            doc.Close()  # 关闭原来的word文件
            print(file + '转换成功！')
            i += 1
        except:
            print(file + '转换失败！')
            files.append(file)  # 将文件名添加到files列表中重新读取
            pass

    print("转换文件%i个" % i)
    word.Quit()
