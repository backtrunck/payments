import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import mysql.connector

class CityPayments():
    def __init__(self, name_city):
        self.name_city = name_city
        user = os.environ.get('USER_DB_PAGAMENTOS')
        host = 'localhost'
        password = os.environ.get('PASS_DB_PAGAMENTOS')
        banco = 'pagamentos'
        try:
            self.conexao = mysql.connector.connect(user = user,  password = password,  host = host,  database = banco)
        except Exception as e:
            print(e)
            return None
    def get_most_payment(self, year_from = '', fonte_recursos = [], recurso_federal = False, somente_empresas = True):    
        sql_where = ''
        sql = '''select cnpj_cpf,nome_credor, sum(vl_bruto) as valor_bruto ,
                        sum(vl_bruto_federal) as valor_bruto_federal
                        from pagamentos_credor
                        {}
                        group by cnpj_cpf,nome_credor
                        order by sum(vl_bruto)	desc ''' 
#        where vl_bruto_federal > 0 and 
#        length(cnpj_cpf) > 11 and exercicio_pagamento > 2016 and 
#        cd_fonte_recursos >= 95
        if year_from:
            sql_where = self.__include_and_where(sql_where, ' exercicio_pagamento > ' + year_from)
            
        if fonte_recursos:
            clausule = '(' + ' or '.join(['cd_fonte_recursos = {}'.format(fonte) for fonte in fonte_recursos]) + ')'
            sql_where = self.__include_and_where(sql_where, clausule)
            
        if recurso_federal:
            sql_where = self.__include_and_where(sql_where, ' vl_bruto_federal > 0 ')
            
        if somente_empresas:
            sql_where = self.__include_and_where(sql_where, ' length(cnpj_cpf) > 11 ')
        
        if sql_where :
            sql_where = ' where municipio = \'' + self.name_city + '\' and ' + sql_where
        else:
            sql_where = ' where municipio = \'' + self.name_city + '\' '
        sql = sql.format(sql_where)
        
        return self.__get_data(sql)
        
    def __include_and_where(self, sql_where, clausule):        
        if sql_where:
            sql_where += ' and ' + clausule
        else:
            sql_where = clausule
        return sql_where
        
    def __get_data(self, sql):
        try:
            cursor = self.conexao.cursor()
            cursor.execute(sql)
            return cursor
            
        except Exception as e:
            print(e)
            return None
def main():
    city_payments = CityPayments('São Félix do Coribe')
    
    row_delta = 0 
    column_delta = 0 
    workbook = xlsxwriter.Workbook('demo.xlsx')
    
    worksheet = workbook.add_worksheet('pgtos_x_fornec')
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    payments = city_payments.get_most_payment(year_from = '', fonte_recursos = [], recurso_federal = False, somente_empresas = False)    
    write_in_sheet(worksheet, row_delta, column_delta, payments, workbook)
    
    worksheet = workbook.add_worksheet('pgtos_x_fornec PJ')
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    payments = city_payments.get_most_payment(year_from = '', fonte_recursos = [], recurso_federal = False, somente_empresas = True)    
    write_in_sheet(worksheet, row_delta, column_delta, payments, workbook)
    
    worksheet = workbook.add_worksheet('pgtos_x_fornec_Rec_Fed')    
    payments = city_payments.get_most_payment(year_from = '', fonte_recursos = [], recurso_federal = True, somente_empresas = True)    
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    write_in_sheet(worksheet, row_delta, column_delta, payments, workbook)
    
    worksheet = workbook.add_worksheet('pgtos_x_fornec_Rec_Fed_2017')
    payments = city_payments.get_most_payment(year_from = '2016', fonte_recursos = [], recurso_federal = True, somente_empresas = True)
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    write_in_sheet(worksheet, row_delta, column_delta, payments, workbook)
    
    worksheet = workbook.add_worksheet('pgtos_precatorios_2017')
    payments = city_payments.get_most_payment(year_from = '2016', fonte_recursos = [95, 96], recurso_federal = False, somente_empresas = True)
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    write_in_sheet(worksheet, row_delta, column_delta, payments, workbook)
    
    worksheet = workbook.add_worksheet('pgtos_precat_FUNDEF_2017')
    payments = city_payments.get_most_payment(year_from = '2016', fonte_recursos = [95], recurso_federal = False, somente_empresas = True)
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 60)
    worksheet.set_column(2, 3, 20)
    write_in_sheet(worksheet, row_delta, column_delta, payments, workbook)
    
    workbook.close()


def write_in_sheet(sheet, row_delta, column_delta, payments, workbook):
    if payments:
        row = 0 
        bold = workbook.add_format({'bold': True})

        #escreve cabeçalho
        for column, value in enumerate( payments.column_names):
            sheet.write(row + row_delta, column + column_delta, value, bold)
        
        data = payments.fetchall() 
        
        for row , data_row in enumerate(data):
            for column, value in enumerate(data_row):
                sheet.write(row + 1 + row_delta, column + column_delta, value)
        
        money_bold = workbook.add_format({'num_format': '$#,##0.00','bold': True})
    
        sheet.write(row + 2 + row_delta,\
                    column_delta,\
                    'Total', bold)
        fisrt_cell = xl_rowcol_to_cell(row_delta + 1, column_delta + len(payments.column_names) - 1)
        last_cell = xl_rowcol_to_cell(row + 1 + row_delta, column_delta + len(payments.column_names) - 1)
        sheet.write(row + 2 + row_delta, \
                    column_delta + len(payments.column_names) - 1,\
                    '= SUM( {}:{} )'.format(fisrt_cell, last_cell),\
                    money_bold)
    
        
        payments.close()
    
if __name__ == '__main__':
    main()
