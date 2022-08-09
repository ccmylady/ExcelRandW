# 未完
# 文件打开错误保护
# 信息写入前比较保护
# 序号唯一性检索
# 格式保护
# 默认输入
# 行数限制

import logging
logging.basicConfig(level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
logging.info('程序开始')
logging.info('数据汇集工具，第1版，20220808。\n'
             '表格某些项汇总填入总表,读取所有文件后1次保存模式\n'
             '注意事项：\n'
             '1) 数据源文件必须是xls格式，转化后写入文件为xls格式\n'
             '2) 文件存放于C:\EXCELzz\combined_origin \n'
             '3) 汇总后文件存放于C:\EXCELzz\combined_combined \n')

import xlwt
import xlrd
import os
import time

yes_no_combined=input("请输入字母'y'开始转换:")

if yes_no_combined in ['y', 'Y']:
    # 定义文件路径
    files_orgin_path=r'C:\EXCELzz\combined_origin'
    files_summary_path=r'C:\EXCELzz\combined_combined'

    # 检索origin文件夹下文件
    filenames_tobecombined = os.listdir(files_orgin_path)
    logging.debug(filenames_tobecombined)
    # 筛除非xls格式文件
    for i,filename in enumerate(filenames_tobecombined):
        if os.path.splitext(filename)[1] not in ['.XLS','.xls']:
            filenames_tobecombined.pop(i)
    logging.info("文件夹下共有{}个'.xls'文件需要转换,分别为{}".format(len(filenames_tobecombined),filenames_tobecombined))
    # 为了提示信息显示顺序正确
    time.sleep(0.1)
    # TODO 读取待合并文件列表
    # 测试临时用文件
    # filenames_tobecombined=['附表1、一般纳税人小微企业印花税申报应享未享“六税两费”减免政策（需反馈）(11所).xls',
    #        '附表1、一般纳税人小微企业印花税申报应享未享“六税两费”减免政策（需反馈）(十五所).xls']
    flag_is_tablehead_exist=1
    # 请输入信息所在表名
    sheetname_input = input("\n请输入信息所在表名(如默认'sheet1'请直接按回车):\n")
    if len(sheetname_input) == 0:
        sheetname_input = 'sheet1'
    logging.info("信息所在表名:{}".format(sheetname_input))
    # 为了提示信息显示顺序正确
    time.sleep(0.1)
    # 请输入汇总表表名
    files_summary_name = input("\n请输入汇总表名(如默认'summary.xls'请直接按回车):\n")
    if len(files_summary_name) == 0:
        files_summary_name = 'summary.xls'
    files_summary_fullpath = os.path.join(files_summary_path, files_summary_name)
    logging.info("汇总表名:{}".format(files_summary_fullpath))


    # TODO 读取待合并文件中的信息
    contents_update=[]
    files_num=0
    for file in filenames_tobecombined:
        file_origin_fullpath=os.path.join(files_orgin_path,file)
        #打开源文件
        workbook_r = xlrd.open_workbook(filename=file_origin_fullpath)
        worksheet_r = workbook_r.sheet_by_name(sheet_name=sheetname_input)
        rows = worksheet_r.nrows

        #读取表单中的每一行
        for rowNum1 in range(rows):
            if flag_is_tablehead_exist:
                if files_num>0 and rowNum1==0:
                    continue
            worksheet_list = worksheet_r.row_values(rowx=rowNum1, start_colx=0, end_colx=None)
            # logging.debug(worksheet_list)
            contents_update.append(worksheet_list)
        files_num+=1
    logging.debug(contents_update)

    # TODO 写入合并后文件
    # 创建一个工作簿
    workbook_w = xlwt.Workbook(encoding='utf-8')
    # 创建一个sheet对象,第二个参数是指单元格是否允许重设置，默认为False
    workbook_w_sheet = workbook_w.add_sheet('sheet1combined', cell_overwrite_ok=False)

    for row,rowvaluelist in enumerate(contents_update):
        for column,cellvalue in enumerate(rowvaluelist):
            workbook_w_sheet.write(row,column,cellvalue)

    # 保存文件
    workbook_w.save(files_summary_fullpath)

    logging.info('程序结束')
    # 为了提示信息显示顺序正确
    time.sleep(0.1)
    input('请输入任意键退出\n')