import sys, os, csv
import util
from bs4 import BeautifulSoup

def parse_file_tcm(file_name, encoding='utf-8'):
    file_tcm = open(file_name,'r', encoding = encoding)
    
    dirname = os.path.dirname(file_name)
    
    csv_file_name = util.obter_nome_arquivo_e_extensao(file_name)[0]
    csv_file_name = csv_file_name + '.csv'
    csv_file_name = os.path.join(dirname, csv_file_name)
    
    csv_file = open(csv_file_name,'w', encoding = 'utf-8')
    
    csv_writer = csv.writer(csv_file, delimiter = ';',  quotechar = '"', quoting = csv.QUOTE_ALL)
    
    soup_file_tcm = BeautifulSoup(file_tcm, 'lxml')
    #MUNICIPIO
    select = soup_file_tcm.find('select', {'name': 'ctl00$ContentPlaceHolder1$UnidadeAno1$ddlMunicipio'})
    
    if select:
        csv_writer.writerow([   'CD_MUNICIPIO', \
                                'NOME_MUNICIPIO'])
        options = select.findAll('option')
        if options:
            i = 0 
            j = 0
            for option in options: 
                if j < len(option['value']):
                    j = len(option['value'])
                    
                if i < len(option.get_text()):
                    i = len(option.get_text())
                    
                csv_writer.writerow([   option['value'], \
                                        option.get_text()])
            print('codigo Uni. Municipio Max: ', j, 'Desc Max: ', i)        
            options = None  
        select = None 
    #FONTE
    select = soup_file_tcm.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddlFonte'})
    if select:
        csv_writer.writerow([   'CD_FONTE', \
                                'DS_FONTE'])
        options = select.findAll('option')
        if options:
            i = 0 
            j = 0
            for option in options: 
                if j < len(option['value']):
                    j = len(option['value'])
                    
                if i < len(option.get_text()):
                    i = len(option.get_text())
                csv_writer.writerow([   option['value'], \
                                        option.get_text()])
              
            print('codigo Fonte Max: ', j, 'Desc Max: ', i) 
            options = None  
        select = None 
    #ORGÃOS
    select = soup_file_tcm.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddlOrgao'})
    if select:
        csv_writer.writerow([   'CD_ORGAO', \
                                'DS_ORGAO'])
        options = select.findAll('option')
        if options:
            i =0 
            j = 0
            for option in options: 
                
                if j < len(option['value']):
                    j = len(option['value'])
                    
                if i < len(option.get_text()):
                    i = len(option.get_text())
                csv_writer.writerow([   option['value'], \
                                        option.get_text()])
              
            print('codigo Uni. Orgão Max: ', j, 'Desc Max: ', i)  
            options = None  
        select = None 
        
    #UNIDADE ORÇAMENTÁRIA
    select = soup_file_tcm.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddlUnidadeOrc'})
    if select:
        csv_writer.writerow([   'CD_UN_ORCAMENT', \
                                'DS_UN_ORCAMENT'])
        options = select.findAll('option')
        if options:
            i =0 
            j = 0
            for option in options: 
                if j < len(option['value']):
                    j = len(option['value'])
                    
                if i < len(option.get_text()):
                    i = len(option.get_text())
                    
                csv_writer.writerow([   option['value'], \
                                        option.get_text().split('-')[-1].strip()])
              
            print('codigo Uni. Orç Max: ', j, 'Desc Max: ', i)    
            options = None  
        select = None 
    #FUNÇÃO
    select = soup_file_tcm.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddlFuncao'})
    if select:
        csv_writer.writerow([   'CD_FUNCAO', \
                                'DS_FUNCAO'])
        options = select.findAll('option')
        if options:
            i =0 
            j = 0
            for option in options: 
                if j < len(option['value']):
                    j = len(option['value'])
                    
                if i < len(option.get_text()):
                    i = len(option.get_text())
                
                csv_writer.writerow([   option['value'], \
                                        option.get_text()])
              
            print('codigo Função Max: ', j, 'Desc Max: ', i)  
            options = None  
        select = None 
    
    #ELEMENTO DA DESPESA
    select = soup_file_tcm.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddlElemento'})
    if select:
        csv_writer.writerow([   'CD_ELEM_DESP', \
                                'DS_ELEM_DESP'])
        options = select.findAll('option')
        if options:
            i =0 
            j=0
            for option in options: 
                if j < len(option['value']):
                    j = len(option['value'])
                    
                if i < len(option.get_text()):
                    i = len(option.get_text())
                
                csv_writer.writerow([   option['value'], \
                                        option.get_text().split('-')[-1].strip()])
              
            print('codigo Elemento Max: ', j, 'Desc Max: ', i)  
            options = None  
        
    csv_file.close()
    file_tcm.close()
def main():
    print(len(sys.argv))
    if len(sys.argv) < 2:
        print('Execute "python3 extrair_codigos_tcm arquivo_com_codigos"')
    else:
        if os.path.isfile(sys.argv[1]):
            parse_file_tcm(sys.argv[1])
        else:
            print('Arquivo inválido ou inexistente')
        
if __name__ == '__main__':
    main()
