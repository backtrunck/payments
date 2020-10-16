import os, logging
import mysql.connector
from tkinter import *
from tkinter.filedialog import askopenfilename

import util
from interfaces_graficas import ScrolledText
from payment import ArquivoPagamentoMacroEscritor, \
                    ArquivoPagamentoTcmBaLeitor, \
                    ArquivoPagamentosMacroLeitor

def make_window():
    root = Tk()
    root.title('Importar Pagamentos')
    root.geometry('800x180')
    frm = Frame(root)
    frm.pack(fill = X)
    
    frmLeft = Frame(frm)
    frmLeft.pack(side=LEFT)
    
    frmRight = Frame(frm)
    frmRight.pack(side=RIGHT)
    
    lbl = LabelFrame(frmLeft, text = "Enviar saída para:")    
    lbl.pack(fill=X)
    option = StringVar(root)
    option.set('Arquivo *.csv')
    
    optMenu = OptionMenu(lbl,  option,  'B. de Dados Arq. TCM',  'B. de Dados Arq. Macro',  'Arquivo *.csv')
    optMenu.pack(fill = X, padx=3, pady=3)
    
    scroll = ScrolledText(frm, height = 8)
    Button(frmLeft, text='Importar', command = (lambda p1 = option, p2 = scroll:convert_payment(p1, p2))).pack(fill = X, padx=3, pady=3)
    Button(frmLeft, text='Importar em Lote', command = send_payments).pack(fill = X, padx=3, pady=3)
    
    scroll.pack(fill=X,padx=3, pady=3)
    
    Button(root, text='Sair', command = root.quit).pack(side= BOTTOM, fill = X, padx=3, pady=3)
    root.mainloop()
    
def convert_payment(option, scroll):
    
    if option.get() == 'B. de Dados Arq. Macro':
        
        payment_file_name = askopenfilename(  title='Selecione o arquivo', 
                                        filetypes=(('CVS files', '*.csv'), ('All files', '*')))
        if len(payment_file_name) == 0:
            
            return
        
        payment_file_tcm = ArquivoPagamentosMacroLeitor(payment_file_name, versao_formato = 2)
        
        dirname = os.path.dirname(payment_file_name)
        
    else:
        
        payment_file_name = askopenfilename(  title='Selecione o arquivo', 
                                        filetypes=(('Xls files', '*.xls'), ('All files', '*')))
                                        
        if len(payment_file_name) == 0:
            return
        
        payment_file_tcm =   ArquivoPagamentoTcmBaLeitor(payment_file_name) 
        
        dirname = os.path.dirname(payment_file_name)
        
        #retira a extensão do nome do arquivo
        csv_file_name = util.obter_nome_arquivo_e_extensao(payment_file_name)[0]
        
        csv_file_name = os.path.basename(csv_file_name) + '.csv'
        csv_file_name = os.path.join(dirname, csv_file_name)
        
    qt_pagamentos_inseridos = 0
    vl_pagamentos_inseridos = 0.0
    qt_pagamentos_rejeitados = 0
    vl_pagamentos_rejeitados = 0.0
    qt_credores = 0
    
    scroll.erase()    
    
    if option.get() == 'Arquivo *.csv':    
        
        scroll.settext('Abrindo Arquivo:\n')
        scroll.settext(csv_file_name)
        indice_fixo = scroll.index(INSERT)        
        scroll.settext('\n')
        
        output_file = ArquivoPagamentoMacroEscritor(csv_file_name, versao_formato=0, encoding = 'utf-8')
        output_file.escrever_cabecalho()
        
        msg =   '\nPagamentos Inseridos {}'
        
        for payment in payment_file_tcm:
            output_file.escrever_pagamento(payment)
            
            qt_pagamentos_inseridos += 1
            
            vl_pagamentos_inseridos += payment.valor_bruto
            
            scroll.erase(indice_fixo, END) 
            
            scroll.settext( text =  msg.format( qt_pagamentos_inseridos), \
                                    posicao = indice_fixo) 
        scroll.settext('\nValor Total dos Pagamentos: ' + util.formatar_moeda(vl_pagamentos_inseridos))    
        scroll.settext('\nFinalizado com Sucesso')    
        
    else:
        #montando o nome do arquivo log
        log_file_name = util.obter_nome_arquivo_e_extensao(payment_file_name)[0]
    
        log_file_name = os.path.basename(log_file_name) + '.log'
        log_file_name = os.path.join(dirname, log_file_name)
        
        logging.basicConfig(   level       = logging.INFO, \
                                filename    = log_file_name, \
                                format      = '%(levelname)s;%(asctime)s;%(name)s;%(funcName)s;%(message)s')
            
        log = logging.getLogger(__name__)  
        
        user = os.environ.get('USER_DB_PAGAMENTOS')
        host = 'localhost'
        password = os.environ.get('PASS_DB_PAGAMENTOS')
        banco = 'pagamentos'
        
        
        scroll.settext('Abrindo Conexão\n')
        
        log.info('Abrindo conexão com banco de dados \'{}\' em {} usando usuário {}'.format(banco, host, user))
        
        try:
            conexao = mysql.connector.connect(user = user,  password = password,  host = host,  database = banco)
        except Exception as e:
            log.info('Erro: {}'.format(e))
            scroll.settext('Erro abrindo Conexão\n')
            raise e
        
        log.info('Conexão Aberta com sucesso')
        scroll.settext('Conexão Aberta\n')
        indice_fixo = scroll.index(INSERT)
        indx, indy  = scroll.getindex(INSERT)
        
        msg =   '\nPagamentos Inseridos {}\n' + \
                'Pagamentos Rejeitados {}\n' + \
                'Credores Inseridos {}\n'
        try:
            cursor = conexao.cursor()
        
            for payment in payment_file_tcm:

                qt_credores += inserir_credor(payment.cpf_cnpj, payment.credor, cursor, log)    
                
                sql = "INSERT INTO pagamentos_macro                             \
                            (   municipio,unidade, exercicio_pagamento, empenho,\
                                dotacao, processo, credor, cnpj_cpf,            \
                                data_empenho, data_pagamento, valor_pagamento,  \
                                valor_retencao, valor_bruto, conta_nome,        \
                                banco, agencia, conta, documento, rp,           \
                                contrato, licitacao, historico,cd_fonte_recursos,cd_elemento_despesa)                 \
                        Values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                #campo dotacao tem a seguinte estrutura(orgão,unidade orçamentária/função/subfunção/programa/atividadae/proj_ativ_oper/Elemento da despesa/fonte de recursos
                dotacao = payment.dotacao.split('/')
                fonte_recursos = dotacao[-1].strip()
                elemento_despesa = dotacao[-2].strip()
                lista_pagamento = list(payment)
                lista_pagamento.append(fonte_recursos)
                lista_pagamento.append(elemento_despesa)
                cursor.execute(sql, lista_pagamento)
                
                
                log.info('Inserido Pagamento - (empenho/processo/unidade/exercicio_pagamento) ({}/{}/{:%Y-%m-%d}/{:10.2f})'.format(payment.empenho,  \
                                                                                                                            payment.processo,  \
                                                                                                                            payment.data_pagamento,  \
                                                                                                                            payment.valor_bruto))
                qt_pagamentos_inseridos += 1
                vl_pagamentos_inseridos += payment.valor_bruto
                
                scroll.erase(indice_fixo, END) 
                scroll.settext( text = msg.format( qt_pagamentos_inseridos, \
                                            qt_pagamentos_rejeitados, qt_credores) , \
                                posicao = indice_fixo) 

            scroll.settext('Total de Pagamentos Inseridos: ' + util.formatar_moeda(vl_pagamentos_inseridos))    
            scroll.settext('\nTotal de Pagamentos Rejeitados: ' + util.formatar_moeda(vl_pagamentos_rejeitados))    
            scroll.settext('\nTotal de Pagamentos: ' + util.formatar_moeda(vl_pagamentos_inseridos + vl_pagamentos_rejeitados))   
             
                                
        except Exception as e:
            cursor.close()
            conexao.rollback()
            scroll.settext('Erro - Ações Desfeitas - Detalhes do Erro:{}'.format(e))
            log.error('Erro - {}'.format(e))
            raise(e)
        else:
            cursor.close()            
            conexao.commit()
            conexao.close
            scroll.settext('\nFinalizado com Sucesso')
            
def inserir_credor(cnpj, nome_credor, cursor, log):       
    sql =   '   select cnpj_cpf      \
                from credor          \
                where cnpj_cpf = %s'
                    
    cursor.execute(sql, ( cnpj,))
    cursor.fetchall()
                                
    if cursor.rowcount == 0:
        
        sql = "INSERT INTO credor(cnpj_cpf,nome_credor) values(%s,%s)"
    
        cursor.execute(sql, (cnpj, nome_credor))    
    
        log.info('Credor ({}-{} inserido na base'.format(cnpj, nome_credor))
        return 1
    else:
        log.info('Credor ({}-{} consta na base'.format(cnpj, nome_credor))
        return 0

def send_payments():
    pass
    
def main():
    make_window()
    
if __name__ == '__main__':
    main()
