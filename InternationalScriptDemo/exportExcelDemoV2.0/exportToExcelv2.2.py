# -*- coding: utf-8 -*-
from optparse import OptionParser
import os
import codecs
import xlwt

def addParser():
    parser = OptionParser()
    parser.add_option("-f", "--directory",
                      help="original.xls File Path.",
                      dest="directory")

    parser.add_option("-t", "--targetFile",
                      help="Target Floder Path.",
                      dest="target")
    
    # parser.add_option("-i", "--iOSAdditional",
    #                   help="iOS additional info.",
    #                   metavar="iOSAdditional")
    return parser

def readKeysAndValuesFromeFilePath(path):
    '''
        读取文件的键值对
        :param path: 文件路径
        :return: 键值对
    '''
    if path is None:
        return
    listKey = []
    listValue = []
    for string in codecs.open(path,'r','utf-8').readlines():
        list = string.split(' = ')
        if len(list) >= 2:
            listKey.append(list[0].lstrip('"').rstrip('"'))
            listValue.append(list[1].lstrip('"').rstrip('\n').rstrip(';').rstrip('"'))
    return (listKey, listValue)


def exportToExcel(options):
    # 解析参数
    option, args = options.parse_args()
    directory = option.directory    # "iOSLocal"
    targetFile = option.target  # "localizableToExcel.xls"
    if directory is not None:
        index = 0
        # if targetFile is not None:
        wb = xlwt.Workbook()
        ws = wb.add_sheet('test', cell_overwrite_ok=True)

        for parent, dirnames, filenames in os.walk(directory):
            for dirname in dirnames:
                # Key 和 国家简码
                if index == 0:
                    ws.write(0, 0, "Key")
                # xx.proj 取xx 表示本地化国家简码
                countryCode = dirname.split('.')[0]
                ws.write(0, index + 1, countryCode)

                # Key 和value
                path = directory + '/' + dirname + '/Localizable.strings'
                (keys, values) = readKeysAndValuesFromeFilePath(path)

                for x in range(len(keys)):
                    key = keys[x]
                    value = values[x]
                    if (index == 0):
                        ws.write(x + 1, 0, key)
                        ws.write(x + 1, 1, value)
                    else:
                        ws.write(x + 1, index + 1, value)
                index += 1

        wb.save(targetFile)

def main():
    options = addParser()
    exportToExcel(options)
    # exportToExcel()


if __name__ == '__main__':
    main()