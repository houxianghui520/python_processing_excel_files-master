import os
import re
import time

import pandas as pd
import xlwings as xw

ImportFilePath = 'import/'  # 待处理文件目录
ExportFilePath = 'export/'  # 输出文件目录
BaseFilePath = 'base.xlsx'  # 模板文件路径
if not os.path.exists(ImportFilePath):
    os.makedirs(ImportFilePath)
if not os.path.exists(ExportFilePath):
    os.makedirs(ExportFilePath)


def ListDir(path: str) -> list:
    '''输入路径，返回路径下所有文件名'''
    def add_path(filename):
        return path + filename
    list_ = os.listdir(path)
    paths = list(map(add_path, list_))
    return paths


def GetTime() -> str:
    '''返回当前时间字符串,用于命名'''
    return time.strftime('-%Y-%m-%d-%H-%M', time.localtime(time.time()))


def SaveAs(name: str) -> str:
    '''输入目标命名,复制并重命名目标文件,返回路径'''
    os.system('copy %s export/%s' % (BaseFilePath, BaseFilePath))
    if re.findall('rt/(.*?).xls', name, re.S):
        name_ = re.findall('rt/(.*?).xls', name, re.S)[0]
    else:
        name_ = '命名失败'
    ex_name = ExportFilePath + name_  + ' V1.xlsx'
    print(ex_name)
    os.rename('export/%s' % (BaseFilePath), ex_name)
    return ex_name


def PdRead(file_im: str) -> int:
    '''输入文件路径,进行处理,并输出文件output.xls,返回修改后行数'''
    df = pd.read_excel(file_im, skiprows=2, dtype={
        '单位编号': str, '单位名称': str, '报表编号': str, '报表名称': str,
        '检验结果': str, '坐标': str, '报表项目': str, '报表结果': str,
        '校验公式结果': str, '差额': str, '校验公式': str, '单位': str,
        '报表': str, '公式性质': str, '股份坐标': str, '转换公式': str})

    # order = ['报表', '报表名称', '公式性质', '坐标', '报表项目', '报表结果', '差额',
    #          '校验公式', '单位', '单位名称', '单位编号', '报表编号', '校验结果', '校验公式结果']
    # df = df[order]
    # 重新排序列
    # df.drop(['报表编号'], axis=1, inplace=True)
    # # 删除列

    del_list = []
    for i in range(df.shape[0]):
        if DelCow(df, i):
            del_list.append(i)
    df.drop(del_list, inplace=True)
    print('删除成功', del_list)
    # 根据规则删除行

    df = df.reset_index(drop=True)
    # 重置索引
    del_list = []
    for i in range(df.shape[0]-1):
        # 删除已清空行
        if isinstance(df.iat[i, 0], str) and isinstance(df.iat[i+1, 0], str):
            del_list.append(i)
        if isinstance(df.iat[i, 5], str) and isinstance(df.iat[i, 10], str):
            # 删除重复行
            if df.iat[i, 5] == df.iat[i+1, 5] and df.iat[i, 10] == df.iat[i+1, 10]:
                del_list.append(i)
    # if  df.iat[df.shape[0]-1, 1] != '':
    #     del_list.append(df.shape[0]-1)
    df.drop(del_list, inplace=True)
    # 删除已清空项

    # print(df.head())
    # book = load_workbook(file_ex)
    # writer = pd.ExcelWriter(file_ex, engine='openpyxl')
    # writer.book = book
    # df.to_excel(writer, '校验报告-过程', index = False)
    # writer.save()
    # 全sheet导出

    # 给=字符开头的字符串加上'
    df['校验公式'] = ["'%s" % i for i in df['校验公式']]
    df['单位'] = ["'%s" % i for i in df['单位']]
    df['单位编号'] = ["'%s" % i for i in df['单位编号']]

    #替换所有'nan的字符串
    df.replace("'nan", '', inplace=True)
    df.to_excel('output.xlsx', index=False)
    return df.shape[0]


def DelCow(df: object, i: int) -> bool:
    '''输入df和行数, 进行剔除判定, 返回布尔值'''
    if '未通过' not in df.iat[i, 4]:  # 删除所有已通过和未校验
        return True
    if isinstance(df.iat[i, 14], float) and isinstance(df.iat[i, 12], str):
        # 删除无校验公式行
        return True
    if isinstance(df.iat[i, 10], float) and isinstance(df.iat[i, 12], str):
        # 删除无股份坐标行
        return True      
    # if df.iat[i, 9] == 'nan' and isinstance(df.iat[i, 12], str):
    #      #剔除差额为0项
    #     return True 

    if isinstance(df.iat[i, 10], str):
        if isinstance(df.iat[i, 6], str):
            if '变动行' in df.iat[i, 6]:  # 删除变动行
                return True
        if 'BD(' in df.iat[i, 10]:  # 删除变动行
            return True
        if '(-1,' in df.iat[i, 10]:  # 上年同期数问题删除
            return True
        if 'BB(0, -1@,' in df.iat[i, 10]:  # 上月数问题删除
            return True
        if df.iat[i, 12] == 'BA01':
            if df.iat[i, 5][0] in 'DH':  # BA01
                return True
        if df.iat[i, 12] == 'BA02':
            if df.iat[i, 5][0] in 'DHEJ':  # BA01
                return True
        if df.iat[i, 12] == 'BA04':  # BA04
            if df.iat[i, 5] == 'C27' and 'BA08' in df.iat[i, 10]:
                return True
            # if df.iat[i, 5] == 'G31' and 'BA01' in df.iat[i, 10]:
            #     return True
            # if df.iat[i, 5] == 'G32' and 'BA01' in df.iat[i, 10]:
            #     return True
            if df.iat[i, 5][0] in 'DH':
                return True
        if df.iat[i, 12] == 'BA05':  # BA05
            if df.iat[i, 5] == 'C8' and 'BA02' in df.iat[i, 10]:
                return True
            if df.iat[i, 5] == 'C19':# 报错为BD39的c79,递延所得税负债转回的递延所得税费用
                return True
            if df.iat[i, 5] == 'C23':
                return True
            # if df.iat[i, 5] == 'C32' and 'BA01' in df.iat[i, 10]:
            #     return True
            # if df.iat[i, 5] == 'C33' and 'BA01' in df.iat[i, 10]:
            #     return True
            if df.iat[i, 5][0] in 'DH':
                return True
        if df.iat[i, 12] == 'BA08':  # BA08, 境外的才可忽略
            # if '>=' in df.iat[i, 10]:
            #     return True
            if '=0' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BA10':  # BA10
            if df.iat[i, 5] == 'C20':#因同时出现经营累积和经营减值而报错, 删除
                return True

        if df.iat[i, 12] == 'BB01':
            if df.iat[i, 5] in  ['G22','K31','K37']:
                return True
            if df.iat[i, 5] == 'G34':
                return True
            if df.iat[i, 5][0] in 'DHL':  # BB01
                return True
        if df.iat[i, 12] == 'BB02':
            if df.iat[i, 5][0] in 'CDEF':
                return True
            if df.iat[i, 5] in ['H23', 'H24', 'G17', 'H17']:
                return True
        if df.iat[i, 12] == 'BB03':
            if df.iat[i, 5][0] in 'EFJ':
                return True
            if df.iat[i, 5] == 'C8' and '=BB(BB02, G21:G21)-BB(BB02, G40:G40)-BB(BB02, G28:G28)' in df.iat[i, 10]:
                return True
            if df.iat[i, 5] == 'E8' and '=BB(BB02, N21:N21)-BB(BB02, N40:N40)-BB(BB02, N28:N28)' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BB04':
            if df.iat[i, 5][0] in 'DH':
                return True
            if 'ZC01' in df.iat[i, 10]:
                return True
            if df.iat[i, 5] == 'C21':
                return True
        if df.iat[i, 12] == 'BB10':
            if df.iat[i, 5][0] in 'DH':
                return True
            if df.iat[i, 5] == 'C18':
                return True
        if df.iat[i, 12] == 'BB11':
            if df.iat[i, 5][0] in 'DH':
                return True
            if df.iat[i, 5] in ['G17','G10']:
                return True
        if df.iat[i, 12] == 'BD01':
            if df.iat[i, 5][0] == 'R':
                return True
            if ':R' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BD01-1':
            if df.iat[i, 5][0] in 'EHIJKL':
                return True
        if df.iat[i, 12] == 'BD02':
            if df.iat[i, 5][0] in 'CDE':
                return True
        if df.iat[i, 12] == 'BD03':
            if df.iat[i, 5][:2] in 'AHAIAJ':
                return True
            if df.iat[i, 5][0] == 'C':
                return True
            if df.iat[i, 5] == 'AH24':
                return True
            if df.iat[i, 5] in ['R12', 'R19', 'R30'] and '<=' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BD07':
            if df.iat[i, 5][0] in 'EFGJN':
                return True
        if df.iat[i, 12] == 'BD10':
            if df.iat[i, 5] == 'Q8':
                return True
        if df.iat[i, 12] == 'BD10-1':
            if df.iat[i, 5][-1:] == '7':
                return True
        if df.iat[i, 12] == 'BD14':
            if df.iat[i, 5][0] in 'CE':
                return True
            if df.iat[i, 5] in ['D40','E40']:
                return True        
        if df.iat[i, 12] == 'BD19':
            if df.iat[i, 5] == 'F47':
                return True
        if df.iat[i, 12] == 'BD22':
            if 'BA09' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BD24':
            if df.iat[i, 5][0] in 'DFH':
                return True
        if df.iat[i, 12] == 'BD26':
            if df.iat[i, 5][0] in 'IJ':
                return True
        if df.iat[i, 12] == 'BD32':
            if df.iat[i, 5][0] in 'CH':
                return True
        if df.iat[i, 12] == 'BD33':
            if df.iat[i, 5][0] in 'EJ':
                return True    
            if df.iat[i, 5] == 'I52':
                return True    
        if df.iat[i, 12] == 'BD34':
            if df.iat[i, 5][0] in 'E':
                return True
            if df.iat[i, 5] == 'D24':
                return True                      
        if df.iat[i, 12] == 'BD35':
            if df.iat[i, 10] == '=0':
                return True
        if df.iat[i, 12] == 'BD36':
            if df.iat[i, 5][0] in 'E':
                return True    
        if df.iat[i, 12] == 'BD37':
            if df.iat[i, 5][0] in 'E':
                return True    
        if df.iat[i, 12] == 'BD38':
            if df.iat[i, 5][0] in 'CEOGHMN':
                return True
            if df.iat[i, 5][-2:] in ['13', '21','41','44','54']:
                return True
            if df.iat[i, 5] in ['C39', 'D39', 'AE54','T54']:
                return True
        if df.iat[i, 12] == 'BD39':
            if df.iat[i, 5][0] in 'D':
                return True
            if df.iat[i, 5] in ['C81','C97']:
                return True    
            if df.iat[i, 5][-2:] not in [str(i) for i in range(81, 104)]:
                return True
        if df.iat[i, 12] == 'BD46':
            if df.iat[i, 5][0] in 'FGH':
                return True
        if df.iat[i, 12] == 'BD46-1':
            if df.iat[i, 5][0] in 'CDF':
                return True
        if df.iat[i, 12] == 'BD47':
            if df.iat[i, 5][0] in 'CEGI':
                return True
        if df.iat[i, 12] == 'BD48':
            return True
        if df.iat[i, 12] == 'BD54' and df.iat[i, 5][0] == 'C':
            return True
        if df.iat[i, 12] == 'BY01':
            if df.iat[i, 5][0] in 'DH':
                return True
        if df.iat[i, 12] == 'BY09':
            if df.iat[i, 5][0] in 'D':
                return True

        # 以下为需手工确认项目
        if df.iat[i, 12] == 'BA01' and  df.iat[i, 5] in ['C59','G22','G31','H26']:
            return True
        if df.iat[i,12] == 'BD10' and df.iat[i, 5] in ['N14','X14']:
            return True
        if df.iat[i,12] == 'BD26' and df.iat[i, 5] in ['C17','M17']:
            return True
        if df.iat[i, 12] == 'BA04':  # BA04
            if df.iat[i, 5] == 'G31' and 'BA01' in df.iat[i, 10]:
                return True
            if df.iat[i, 5] == 'G32' and 'BA01' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BA05':  # BA05
            if df.iat[i, 5] == 'C32' and 'BA01' in df.iat[i, 10]:
                return True
            if df.iat[i, 5] == 'C33' and 'BA01' in df.iat[i, 10]:
                return True
        if df.iat[i, 12] == 'BD03':
            if df.iat[i, 5][:2] in 'AG':
                return True
        if df.iat[i, 12] == 'BA08':  # BA08, 境外的才可忽略
            if '>=' in df.iat[i, 10]:
                return True
        #以下为境外投资企业报表
        if df.iat[i, 12][:2] == 'BY':
            return True

    return False


def Merge(path_im: str, path_ex: str, lines: int):
    '''输入pandas整理后文件路径,输出文件路径,电子表格行数'''
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    wb1 = app.books.open(path_im)
    sht1 = wb1.sheets(1)
    # lastcell = sht1.range('A2').expand().last_cell.row  # 最后一行
    # print(lastcell)
    rg = "A2:P" + str(lines)
    data = sht1.range(rg).value  # 获取范围数据
    wb1.close()
    # print(data)

    wb2 = app.books.open(path_ex)
    sht2 = wb2.sheets(1)
    # i = 0
    # for d in data:
    #     cell = 'A'+str(i+4)
    #     sht2.range(cell).value = d
    #     i += 1
    sht2.range('A4').value = data
    wb2.save()
    app.quit()


if __name__ == "__main__":
    files = ListDir(ImportFilePath)  # 所有待处理文件,list
    print(files)
    for file_ in files:
        file_ex = SaveAs(file_)  # 模板另存路径,str
        lines = PdRead(file_)  # 经过pandas处理后的文件行数
        Merge(path_im='output.xlsx', path_ex=file_ex, lines=lines)
    print('******处理完成*****')
    input('按任意键退出')
