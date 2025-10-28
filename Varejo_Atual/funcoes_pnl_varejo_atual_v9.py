import numpy as np
import datetime
from datetime import datetime
import pandas as pd

pd.options.display.float_format = '{:,.2f}'.format

def retirar_ponto_zero(row, col):
    '''Retira ponto e zero de valores que deveriam ser números, varrendo linha a linha do DataFrame, passar como lambda function.
    
    Parameters:
    row(linha): Linha do DataFrame
    col(str): String com o nome da Coluna que passará por tratamento
    
    Returns:
    row[anomes](list): coluna com linhas ajustadas
    ''' 
    if '.' in row[col]:
        row[col] = row[col][:-2]
    else:
        row[col] = row[col]
    return row[col]

def gerar_df_pnl(cmodelo_fin, cenario, anomes_m0, anomes_m1):

    # linhas que compõem o P&L
    linhas_cascada = ['12', '13', '17', '40', '72', '34', '83', '69', '70', '92', '29', '93', '102', '107', '108', '111']
    # ordem das linhas no P&L
    sorter = ['Margem', 'Crédito', 'Captação', 'Demais Margens', 'Comissões', 'Alocação de Capital', 'ROF/Orex/Equiv./Div.', 'MOB', 'PDD', 'MOL', 'Gastos', 'Oryp / Outros Ativos', 'BAI']
    
    m0 = datetime.now().month
    
    # Caso seja Avance/Fehcamento/Prévia de Janeiro
    
    if m0 == 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27') or m0 == 1 and cenario in ('P27','Negócios'):
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '12'
        m1 = '11'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y0}{m1}'
        ANOMES_M0_LY = f'{y1}{m0}'
        
    elif m0 == 1 and cenario.lower() in ('previa', 'prévia', 'prev', 'prév'):
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '01'
        m1 = '12'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y1}{m1}'
        ANOMES_M0_LY = f'{y1}{m0}'
        
    # Caso seja Avance/Fechamento de Fevereiro        
        
    elif m0 == 2 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27') or m0 == 2 and cenario in ('P27','Negócios'):
        y0 =  datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '01'
        m1 = '12'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y1}{m1}'
        ANOMES_M0_LY = f'{y1}{m0}'
        
    # Demais meses de Avance/Fechamento
        
    elif (m0 != 1 and m0 != 2) and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27') or (m0 != 1 and m0 != 2) and cenario in ('P27','Negócios'):
        y0 =  datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = m0-1
        if len(str(m0)) == 1:
            m0 = '0' + str(m0)
        m1 = int(m0)-1
        if len(str(m1)) == 1:
            m1 = '0' + str(m1)      
        m0 = str(m0)
        m1 = str(m1)
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y1}{m1}'
        ANOMES_M0_LY = f'{y1}{m0}' 
        
    # Caso seja P26/27/28 ...
        
        
    # Demais meses e cenários de Prévia
        
    else:
        y0 =  datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m1 = m0 - 1
        if len(str(m0)) == 1:
            m0 = '0' + str(m0)
        if len(str(m1)) == 1:
            m1 = '0' + str(m1)
        m0 = str(m0)
        m1 = str(m1)
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y0}{m1}'
        ANOMES_M0_LY = f'{y1}{m0}'
        
        
    # colunas para somar YTD
    mes = ['01','02','03','04','05','06','07','08','09','10','11','12']
    ## DECLARAÇÃO DE VARIAVEIS
    cont = 0 ## incluido léo santos 
    m_ref = int(m0)
    FY_mes = []
    YTD_y1 = []
    YTD_y0 = []
    for i in range(len(mes)):
        if i == m_ref:
            cont = i
            break
    FY_mes = mes[:cont]
    
    for j in FY_mes:
        YTD_y1.append(f'{y1}{j}')
        YTD_y0.append(f'{y0}{j}')
        
    # meses de comparação
    dict_mes = {
    '01':'Jan',
    '02':'Fev',
    '03':'Mar',
    '04':'Abr',
    '05':'Mai',
    '06':'Jun',
    '07':'Jul',
    '08':'Ago',
    '09':'Set',
    '10':'Out',
    '11':'Nov',
    '12':'Dez',
    '00':'Dez'
    }

    header_M0 = dict_mes.get(ANOMES_M0[-2:]) + '/' + str(ANOMES_M0[2:4])
    
    for i in range(len(cmodelo_fin)):
        cmodelo_fin[i]['Cod'] = cmodelo_fin[i]['Cod'].astype(str)
        cmodelo_fin[i]['Cod'] = cmodelo_fin[i].apply(lambda row: retirar_ponto_zero(row,'Cod'), axis = 1)

    ## CASCADA VAREJO AMPLIADO

    # buscas as linhas do cascada que formam o P&L resumido
    pnl_cascada_varejo_ampliado = cmodelo_fin[0].loc[cmodelo_fin[0]['Cod'].isin(linhas_cascada)]
    pnl_cascada_varejo_ampliado['Cod'] = pnl_cascada_varejo_ampliado['Cod'].astype(int)
    pnl_cascada_varejo_ampliado['Cod'] = pnl_cascada_varejo_ampliado['Cod'].astype(str)
   

    # soma de YTD de 2022 e 2023
    pnl_cascada_varejo_ampliado[f'YTD{y1[-2:]}'] = pnl_cascada_varejo_ampliado[YTD_y1].sum(axis=1)
    pnl_cascada_varejo_ampliado[f'YTD{y0[-2:]}'] = pnl_cascada_varejo_ampliado[YTD_y0].sum(axis=1)

    # filtrar colunas necessárias
    df_pnl_VarTotal = pnl_cascada_varejo_ampliado[['Cod','Itens / Período', ANOMES_M0_LY, ANOMES_M1, ANOMES_M0, f'YTD{y1[-2:]}', f'YTD{y0[-2:]}']]

    # divisão da coluna de resultado no mês para R$ Mil
    df_pnl_VarTotal[ANOMES_M0] = df_pnl_VarTotal[ANOMES_M0]/1E3
    df_pnl_VarTotal[ANOMES_M1] = df_pnl_VarTotal[ANOMES_M1]/1E3
    df_pnl_VarTotal[ANOMES_M0_LY] = df_pnl_VarTotal[ANOMES_M0_LY]/1E3
    df_pnl_VarTotal[f'YTD{y1[-2:]}'] = df_pnl_VarTotal[f'YTD{y1[-2:]}']/1E3
    df_pnl_VarTotal[f'YTD{y0[-2:]}'] = df_pnl_VarTotal[f'YTD{y0[-2:]}']/1E3

    # renomeando o P&L 
    df_pnl_VarTotal['Cascada'] = df_pnl_VarTotal['Itens / Período']
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['12']), 'Cascada'] = 'Margem'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['13']), 'Cascada'] = 'Crédito'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['17']), 'Cascada'] = 'Captação'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['40']), 'Cascada'] = 'Comissões'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['34']), 'Cascada'] = 'Alocação de Capital'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['92']), 'Cascada'] = 'MOB'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['29']), 'Cascada'] = 'PDD'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['93']), 'Cascada'] = 'MOL'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['102']), 'Cascada'] = 'Gastos'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['111']), 'Cascada'] = 'BAI'

    # agrupando linhas
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['72', '83', '69', '70']), 'Cascada'] = 'ROF/Orex/Equiv./Div.'
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'].isin(['107', '108']), 'Cascada'] = 'Oryp / Outros Ativos'

    # agrupando por Cascada
    df_pnl_VarTotal_1 = df_pnl_VarTotal[['Cod','Cascada', ANOMES_M0_LY, ANOMES_M1, ANOMES_M0]].groupby('Cascada').sum().reset_index()
    df_pnl_VarTotal_2 = df_pnl_VarTotal[['Cod','Cascada', f'YTD{y1[-2:]}', f'YTD{y0[-2:]}']].groupby('Cascada').sum().reset_index()

    df_pnl_VarTotal = df_pnl_VarTotal_1.merge(df_pnl_VarTotal_2, on='Cascada', how='left')

    # cálculo de "Demais Margens"
    df_pnl_VarTotal_margem = df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'] == '12']
    df_pnl_VarTotal_capt   = df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'] == '17']
    df_pnl_VarTotal_cred   = df_pnl_VarTotal.loc[df_pnl_VarTotal['Cod'] == '13']

    df_pnl_VarTotal_capt[df_pnl_VarTotal_capt.select_dtypes(include=['number']).columns] *= -1
    df_pnl_VarTotal_cred[df_pnl_VarTotal_cred.select_dtypes(include=['number']).columns] *= -1

    df_pnl_VarTotal_union = pd.concat([df_pnl_VarTotal_margem, df_pnl_VarTotal_capt, df_pnl_VarTotal_cred])

    df_pnl_VarTotal_union['Cod'] = 'SOMA'
    df_pnl_VarTotal_union['Cascada'] = 'Demais Margens'
    
    df_pnl_VarTotal_union = df_pnl_VarTotal_union.groupby(['Cascada', 'Cod']).sum().reset_index()
    
    df_pnl_VarTotal = pd.concat([df_pnl_VarTotal, df_pnl_VarTotal_union])

    # ordenando por linhas do cascada
    sorterIndex = dict(zip(sorter, range(len(sorter))))

    df_pnl_VarTotal['Ordem'] = df_pnl_VarTotal['Cascada'].map(sorterIndex)

    df_pnl_VarTotal.sort_values(['Ordem'], ascending = True, inplace = True)
    df_pnl_VarTotal.drop('Ordem', 1, inplace = True)

    # removendo o código de linhas agrupadas
    df_pnl_VarTotal.loc[df_pnl_VarTotal['Cascada'].isin(['ROF/Orex/Equiv./Div.', 'Oryp / Outros Ativos']), 'Cod'] = 'SOMA'

    # renomeando o nome da coluna de ANOMES_M0
    df_pnl_VarTotal = df_pnl_VarTotal[['Cod','Cascada', ANOMES_M0_LY, ANOMES_M1, ANOMES_M0, f'YTD{y1[-2:]}', f'YTD{y0[-2:]}']]
    df_pnl_VarTotal = df_pnl_VarTotal.rename(columns= {ANOMES_M0: header_M0})


    # coluna de MoM
    df_pnl_VarTotal['MoM Abs'] = df_pnl_VarTotal[header_M0] - df_pnl_VarTotal[ANOMES_M1]
    df_pnl_VarTotal['MoM %'] = (df_pnl_VarTotal[header_M0]/df_pnl_VarTotal[ANOMES_M1])-1

    # coluna de YoY
    df_pnl_VarTotal['YoY Abs'] = df_pnl_VarTotal[header_M0] - df_pnl_VarTotal[ANOMES_M0_LY]
    df_pnl_VarTotal['YoY %'] = (df_pnl_VarTotal[header_M0]/df_pnl_VarTotal[ANOMES_M0_LY])-1

    # colna YTD YoY
    df_pnl_VarTotal['YTD YoY Abs'] = df_pnl_VarTotal[f'YTD{y0[-2:]}'] - df_pnl_VarTotal[f'YTD{y1[-2:]}']
    df_pnl_VarTotal['YTD YoY %'] = (df_pnl_VarTotal[f'YTD{y0[-2:]}']/df_pnl_VarTotal[f'YTD{y1[-2:]}'])-1


    # faltam colunas de delta PPTO e delta P26 !!!!

    # formatação final

    df_pnl_VarTotal = df_pnl_VarTotal[['Cod', 'Cascada', header_M0, 'MoM Abs', 'MoM %', 'YoY Abs', 'YoY %', f'YTD{y0[-2:]}', 'YTD YoY Abs', 'YTD YoY %']]
    
    
    
    
    ## CASCADA VAREJO

    # buscas as linhas do cascada que formam o P&L resumido
    pnl_cascada_varejo = cmodelo_fin[4].loc[cmodelo_fin[4]['Cod'].isin(linhas_cascada)]
    pnl_cascada_varejo['Cod'] = pnl_cascada_varejo['Cod'].astype(int)
    pnl_cascada_varejo['Cod'] = pnl_cascada_varejo['Cod'].astype(str)

    # soma de YTD de 2022 e 2023
    pnl_cascada_varejo[f'YTD{y1[-2:]}'] = pnl_cascada_varejo[YTD_y1].sum(axis=1)
    pnl_cascada_varejo[f'YTD{y0[-2:]}'] = pnl_cascada_varejo[YTD_y0].sum(axis=1)

    # filtrar colunas necessárias
    df_pnl_Var = pnl_cascada_varejo[['Cod','Itens / Período', ANOMES_M0_LY, ANOMES_M1, ANOMES_M0, f'YTD{y1[-2:]}', f'YTD{y0[-2:]}']]

    # divisão da coluna de resultado no mês para R$ Mil
    df_pnl_Var[ANOMES_M0] = df_pnl_Var[ANOMES_M0]/1E3
    df_pnl_Var[ANOMES_M1] = df_pnl_Var[ANOMES_M1]/1E3
    df_pnl_Var[ANOMES_M0_LY] = df_pnl_Var[ANOMES_M0_LY]/1E3
    df_pnl_Var[f'YTD{y1[-2:]}'] = df_pnl_Var[f'YTD{y1[-2:]}']/1E3
    df_pnl_Var[f'YTD{y0[-2:]}'] = df_pnl_Var[f'YTD{y0[-2:]}']/1E3

    # renomeando o P&L
    df_pnl_Var['Cascada'] = df_pnl_Var['Itens / Período']
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['12']), 'Cascada'] = 'Margem'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['13']), 'Cascada'] = 'Crédito'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['17']), 'Cascada'] = 'Captação'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['40']), 'Cascada'] = 'Comissões'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['34']), 'Cascada'] = 'Alocação de Capital'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['92']), 'Cascada'] = 'MOB'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['29']), 'Cascada'] = 'PDD'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['93']), 'Cascada'] = 'MOL'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['102']), 'Cascada'] = 'Gastos'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['111']), 'Cascada'] = 'BAI'

    # agrupando linhas
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['72', '83', '69', '70']), 'Cascada'] = 'ROF/Orex/Equiv./Div.'
    df_pnl_Var.loc[df_pnl_Var['Cod'].isin(['107', '108']), 'Cascada'] = 'Oryp / Outros Ativos'

    # agrupando por Cascada
    df_pnl_Var_1 = df_pnl_Var[['Cod','Cascada', ANOMES_M0_LY, ANOMES_M1, ANOMES_M0]].groupby('Cascada').sum().reset_index()
    df_pnl_Var_2 = df_pnl_Var[['Cod','Cascada', f'YTD{y1[-2:]}', f'YTD{y0[-2:]}']].groupby('Cascada').sum().reset_index()

    df_pnl_Var = df_pnl_Var_1.merge(df_pnl_Var_2, on='Cascada', how='left')

    # cálculo de "Demais Margens"
    df_pnl_Var_margem = df_pnl_Var.loc[df_pnl_Var['Cod'] == '12']
    df_pnl_Var_capt   = df_pnl_Var.loc[df_pnl_Var['Cod'] == '17']
    df_pnl_Var_cred   = df_pnl_Var.loc[df_pnl_Var['Cod'] == '13']

    df_pnl_Var_capt[df_pnl_Var_capt.select_dtypes(include=['number']).columns] *= -1
    df_pnl_Var_cred[df_pnl_Var_cred.select_dtypes(include=['number']).columns] *= -1

    df_pnl_Var_union = pd.concat([df_pnl_Var_margem, df_pnl_Var_capt, df_pnl_Var_cred])

    df_pnl_Var_union['Cod'] = 'SOMA'
    df_pnl_Var_union['Cascada'] = 'Demais Margens'

    df_pnl_Var_union = df_pnl_Var_union.groupby(['Cascada', 'Cod']).sum().reset_index()

    df_pnl_Var = pd.concat([df_pnl_Var, df_pnl_Var_union])

    # ordenando por linhas do cascada
    sorterIndex = dict(zip(sorter, range(len(sorter))))

    df_pnl_Var['Ordem'] = df_pnl_Var['Cascada'].map(sorterIndex)

    df_pnl_Var.sort_values(['Ordem'], ascending = True, inplace = True)
    df_pnl_Var.drop('Ordem', 1, inplace = True)

    # removendo o código de linhas agrupadas
    df_pnl_Var.loc[df_pnl_Var['Cascada'].isin(['ROF/Orex/Equiv./Div.', 'Oryp / Outros Ativos']), 'Cod'] = 'SOMA'

    # renomeando o nome da coluna de ANOMES_M0
    df_pnl_Var = df_pnl_Var[['Cod','Cascada', ANOMES_M0_LY, ANOMES_M1, ANOMES_M0, f'YTD{y1[-2:]}', f'YTD{y0[-2:]}']]
    df_pnl_Var = df_pnl_Var.rename(columns= {ANOMES_M0: header_M0})


    # coluna de MoM
    df_pnl_Var['MoM Abs'] = df_pnl_Var[header_M0] - df_pnl_Var[ANOMES_M1]
    df_pnl_Var['MoM %']   = (df_pnl_Var[header_M0]/df_pnl_Var[ANOMES_M1])-1

    # coluna de YoY
    df_pnl_Var['YoY Abs'] = df_pnl_Var[header_M0] - df_pnl_Var[ANOMES_M0_LY]
    df_pnl_Var['YoY %'] = (df_pnl_Var[header_M0]/df_pnl_Var[ANOMES_M0_LY])-1

    # colna YTD YoY
    df_pnl_Var['YTD YoY Abs'] = df_pnl_Var[f'YTD{y0[-2:]}'] - df_pnl_Var[f'YTD{y1[-2:]}']
    df_pnl_Var['YTD YoY %'] = (df_pnl_Var[f'YTD{y0[-2:]}']/df_pnl_Var[f'YTD{y1[-2:]}'])-1


    # faltam colunas de delta PPTO e delta P26 !!!!

    # formatação final

    df_pnl_Var = df_pnl_Var[['Cod', 'Cascada', header_M0, 'MoM Abs', 'MoM %', 'YoY Abs', 'YoY %', f'YTD{y0[-2:]}', 'YTD YoY Abs', 'YTD YoY %']]
    
    
    return df_pnl_VarTotal, df_pnl_Var