import docx2pdf  # pip install docx2pdf --user
import os


def getAllFilesList(filepath):
    """
    获取指定目录下的所有word文件列表
    :param filepath: 指定目录
    :return: 指定目录下的所有word文件列表
    """
    files = []
    for file in os.listdir(filepath):
        if file.endswith(".docx"):
            files.append(filepath + file)

    return files


def convert_single(src_file, dst_file):
    """
    单个文件实现 word->pdf 转换
    :param src_file: 源文件 xxx.docx
    :param dst_file: 目标文件 xxx.pdf
    :return:
    """
    docx2pdf.convert(src_file, dst_file)


def convert_batch(files_list):
    """
    批量转换 word->pdf
    :param files_list:
    :return:
    """
    for file in files_list:
        docx2pdf.convert(file, file.split('.')[0] + '.pdf')
        print(file + '转换成功！')


if __name__ == '__main__':
    input_file = "D:" + os.sep + "sampple" + os.sep + "word2pdf" + os.sep + "data.docx"
    output_file = "D:" + os.sep + "sampple" + os.sep + "word2pdf" + os.sep + "data.pdf"

    filepath = "D:" + os.sep + "sampple" + os.sep + "word2pdf" + os.sep

    # 1.单个文件转换
    # convert_single(input_file, output_file)

    # 2.批量转换
    files = getAllFilesList(filepath)
    convert_batch(files_list=files)
