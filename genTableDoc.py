#!/usr/bin/python
# -*- coding: UTF-8 -*-
from datetime import datetime
from pydbclib import connect as con
import xlwt

'''
@author:    anke
@contact:   anke.wang@foxmail.com
@file:      genTableDoc.py
@time:      2020/5/22 10:00 上午

@desc:      生成mysql数据库元数据文档
'''


class GenTableDoc(object):
    def __init__(self, connect, database):
        self.db = con(f"mysql+pymysql://{connect}")
        self.database = database
        self.tbls = {}
        self.tbl_en = ''
        self.tbl_cn = ''
        self.sheetName = ''
        self.title_table = ['数据库', '表名', '表注释', '类型', '数据量', '创建时间', '更新时间', '表结构']
        self.title_column = ['表名', '字段序号', '字段', '类型', '主键', '字段描述']

    def _sqls(self, id):
        """
        sql语句
        :param id:
                0：查询某数据库的表列表
                1：根据表名查询表结构
        :return:
        """
        if id == 0:
            return f"""
                    select 
                    b.table_schema,b.table_name,b.table_comment,b.table_type,b.table_rows,b.create_time,b.update_time
                    from information_schema.tables b
                    where 1=1
                    and b.table_schema = '{self.database}';
                    """
        elif id == 1:
            return f"""
                    select 
                    b.table_name,
                    a.ordinal_position,a.column_name,a.column_type,a.column_key,a.column_comment
                    from information_schema.columns a
                    left join information_schema.tables b on a.table_schema=b.table_schema and a.table_name=b.table_name
                    where 1=1
                    and b.table_schema = '{self.database}' -- 表所在数据库
                    and b.table_name = '{self.tbl_en}' -- 你要查的表
                    order by a.table_schema,a.table_name,a.ordinal_position;
                    """

    def _writeExcel(self, excelName, data):
        """
        将表列表及表结构写入excel
        :param excelName: excel名字
        :param data: 表列表数据
        :return:
        """
        writebook = xlwt.Workbook()  # 打开一个excel
        sheet = writebook.add_sheet('表列表')  # 在打开的excel中添加一个sheet

        # 调整样式
        style = xlwt.XFStyle()
        # 边框
        borders = xlwt.Borders()  # Create Borders
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        style.borders = borders
        # 对其方式
        alignment1 = xlwt.Alignment()
        alignment2 = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment1.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        alignment1.vert = 0x01
        # 字体
        fontT = xlwt.Font()
        fontF = xlwt.Font()
        fontT.bold = True  # 字体加粗
        fontT.height = 20 * 13  # 字体大小，13为字号，20为衡量单位
        fontF.bold = False

        # 表列表
        for i, lin in enumerate(data):
            self.sheetName = f'tbl{i}'
            # 第一列填充序号
            if i < len(data) - 1:
                sheet.write(i + 1, 0, i + 1, style)  # 序号
            for j, cell in enumerate(lin):
                # 处理时间格式
                if isinstance(cell, datetime):
                    style.num_format_str = 'YYYY/M/D h:mm:ss'
                elif i == 0:
                    # 处理表头
                    style.font = fontT
                    style.alignment = alignment1
                sheet.write(i, j + 1, cell, style)

                # 清除格式
                style.num_format_str = ''
                style.font = fontF
                style.alignment = alignment2

            # 添加表结构超链接
            if i > 0:
                link = f'HYPERLINK("#{self.sheetName}!B1";"查看")'
                sheet.write(i, len(lin) + 1, xlwt.Formula(link), style)

        # 表结构
        for i, (self.tbl_en, self.tbl_cn) in enumerate(self.tbls.items()):
            self.sheetName = f'tbl{i + 1}'
            db_read = self.db.read(self._sqls(1), as_dict=False)
            sheet = writebook.add_sheet(self.sheetName)

            sheet.write_merge(0, 0, 0, 5, self.tbl_en)
            sheet.write_merge(1, 1, 0, 5, self.tbl_cn)  # 开始行,结束行, 开始列, 结束列
            link = 'HYPERLINK("#表列表!B3";"返回")'
            sheet.write_merge(2, 2, 0, 5, xlwt.Formula(link))

            # 表头
            for j, cell in enumerate(self.title_column):
                style.font = fontT
                style.alignment = alignment1
                sheet.write(4, j, cell, style)

            # 清除格式
            style.font = fontF
            style.alignment = alignment2
            # 表数据
            num_l = 5
            for read in db_read.get_all():
                for j, cell in enumerate(read):
                    sheet.write(num_l, j, cell, style)
                num_l += 1
        writebook.save(excelName)  # 一定要记得保存

    def start(self):
        """
        查询表列表
        :return:
        """
        data = []
        data.append(self.title_table)
        lins = self.db.read(self._sqls(0), as_dict=False)
        for lin in lins.get_all():
            self.tbls[lin[1]] = lin[2]
            data.append(lin)
        print(self.tbls)
        self._writeExcel(f'tableSche_{self.database}.xlsx', data)


if __name__ == '__main__':
    GenTableDoc(connect='root:123456@127.0.0.1:3306', database='dangan').start()
