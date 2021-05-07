#!/usr/bin/env python
import sys, os
from payment import tabular_pagamentos_por_empresa_tcm_ba
def main():
    
    if len(sys.argv) !=2:
        print('Chamada inválida')
        print('tcmpgempresa <arquivo_pagamento_tcm_empresa>')
        return


    nome_arquivo  = sys.argv[1]
    if not os.path.exists(nome_arquivo):
        print(f'Não foi possível encontrar o arquivo {nome_arquivo}')
        return

    tabular_pagamentos_por_empresa_tcm_ba(nome_arquivo)
    
if __name__ == '__main__':
    main()
