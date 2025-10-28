from copy import deepcopy

import pandas as pd
from datetime import datetime

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
    
    if m0 == 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27'):
        y0 = datetime.now().year - 1
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
        
    elif m0 == 2 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27'):
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '01'
        m1 = '12'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y1}{m1}'
        ANOMES_M0_LY = f'{y1}{m0}'
        
    # Demais meses de Avance/Fechamento
        
    elif (m0 != 1 and m0 != 2) and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27'):
        y0 = datetime.now().year
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
        y0 = datetime.now().year
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

def gerar_df_analysis(rmodelo_fin, cmodelo_fin, pmodelo_fin, pmodelop_fin, cenario, segmentos, anomes_m0, anomes_m1, anomes_m2):
    
    # atualização de variáveis
    m0 = datetime.now().month
    
    if m0 == 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27'):
        y0 = datetime.now().year - 1
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '12'
        m1 = '11'
        m2 = '10'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y0}{m1}'
        ANOMES_M2 = f'{y0}{m2}'
        
    elif m0 == 2 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham', 'P27'):
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '01'
        m1 = '12'
        m2 = '11'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y1}{m1}'
        ANOMES_M2 = f'{y1}{m2}'
        
    elif m0 == 1 and cenario.lower() in ('previa', 'prévia', 'prev', 'prév'):
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '01'
        m1 = '12'
        m2 = '11'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y1}{m1}'
        ANOMES_M2 = f'{y1}{m2}'
        
    elif m0 == 2 and cenario.lower() in ('previa', 'prévia', 'prev', 'prév'):
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m0 = '02'
        m1 = '01'
        m2 = '12'
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y0}{m1}'
        ANOMES_M2 = f'{y1}{m2}'
        
    else:
        y0 = datetime.now().year
        y1 = y0 - 1
        y0 = str(y0)
        y1 = str(y1)
        m1 = m0 - 1
        m2 = m1 - 1
        if len(str(m0)) == 1:
            m0 = '0' + str(m0)
        if len(str(m1)) == 1:
            m1 = '0' + str(m1)
        if len(str(m2)) == 1:
            m2 = '0' + str(m2)
        ANOMES_M0 = f'{y0}{m0}'
        ANOMES_M1 = f'{y0}{m1}'
        ANOMES_M2 = f'{y0}{m2}'
 
    mes = ['01','02','03','04','05','06','07','08','09','10','11','12']
    ## DECLARAÇÃO DE VARIAVEIS
    
    # meses de FY_y0
    m_ref = int(m0)
    FY_mes = []
    FY_y0 = []
    cont = 0 ## incluido léo santos 
    for i in range(len(mes)):
        if i == m_ref:
            cont = i
            break
    FY_mes = mes[:cont]
    
    for j in FY_mes:
        FY_y0.append(f'{y0}{j}')

    #colunas de análise no R
    colunas_r_casc = ['Cod', 'Segto', 'Sheet', ANOMES_M0]
    
    # lista de segmentos e sheets do Varejo
    segmentos_var = {
        'PF I':'PF1',
        'PF II':'PF2',
        'Van Gogh':'VG',
        'Select':'Select',
        'Empresas I':'Emp1',
        'Empresas II':'Emp2',
        'Empresas III':'Emp3',
        'Empresas MEI':'Emp_MEI',
        'Governos':'Tot_Gov',
        'Universidades':'Tot_Univ',
    }

    # lista de agrupamentos de segmento
    segmentos_VarTotal = ['Tot_RealEstate','Tot_Toro','Tot_Gira','Tot_Var']
    segmentos_Var = ['Tot_PF','Tot_PJ']
    segmentos_PF = ['PF1','PF2','VG','Select']
    segmentos_PJ = ['Emp1','Emp2','Emp3','Emp_MEI','Tot_Gov','Tot_Univ']

    # conceitos que são utilizados para análise de variação
    conceito_R = ['Margem', 'Comissão', 'Dividendos', 'Equivalência', 'FGC', 'Gastos', 'IR/CS', 
                  'Orex','ORyP','Outros Ativos','PDD Gerencial','PDD Risco','PIS/COFINS','Remuneração',
                  'Reparto CF','Reparto Comissão','Reparto Margem','Reparto Orex','Reparto Rof','Reparto SF',
                  'Repasses','Risco País','Rof']
    conceito_P = ['Comercial']
    
    

    
    ## 1. VALIDAÇÃO DE ESTRUTURA

    
    ### 1.1 BAI nas Sheets R vs. Casc
    
    # linhas de BAI da Sheet R

    BAI_cascada_R = pd.DataFrame()

    for i in range(len(list(segmentos.keys()))):

        df = rmodelo_fin[i].loc[rmodelo_fin[i]['Cod'].isin(['1198'])]
        df = df.loc[df['Conceito'].isin(['BAI'])]
        df['Segto'] = list(segmentos.keys())[i]
        df['Sheet'] = list(segmentos.values())[i]
        BAI_cascada_R = BAI_cascada_R.append(df,ignore_index = True)

    BAI_cascada_R['Cod'] = '1198'

    BAI_cascada_R = BAI_cascada_R[colunas_r_casc]

    # Linhas de BAI da Sheet Cascada

    BAI_cascada = pd.DataFrame()

    for i in range(len(list(segmentos.keys()))):

        df = cmodelo_fin[i].loc[cmodelo_fin[i]['Cod'].isin(['111'])]
        df['Segto'] = list(segmentos.keys())[i]
        df['Sheet'] = list(segmentos.values())[i]
        BAI_cascada = BAI_cascada.append(df,ignore_index = True)

    BAI_cascada['Cod'] = '1198'

    BAI_cascada = BAI_cascada[colunas_r_casc] 
    
    # Gerando coluna de totalizadores de Segmento

    BAI_cascada_R.loc[BAI_cascada_R['Sheet'].isin(segmentos_VarTotal), 'Segto_Tot'] = 'Tot_VarTotal' 
    BAI_cascada_R.loc[BAI_cascada_R['Sheet'].isin(segmentos_Var), 'Segto_Tot'] = 'Tot_Var' 
    BAI_cascada_R.loc[BAI_cascada_R['Sheet'].isin(segmentos_PJ), 'Segto_Tot'] = 'Tot_PJ' 
    BAI_cascada_R.loc[BAI_cascada_R['Sheet'].isin(segmentos_PF), 'Segto_Tot'] = 'Tot_PF'
    
    # Joint entre R e Casc para realizar o delta e ver se os dois batem

    BAI_check = BAI_cascada.merge(BAI_cascada_R, on='Sheet', how='left')
    BAI_check['Check_BAI'] = BAI_check[f'{ANOMES_M0}_x'] - BAI_check[f'{ANOMES_M0}_y']
    BAI_check = BAI_check[['Sheet', 'Check_BAI']]
    
    
    
    
    ### 1.2 BAI por Segmento na Sheet R
    
    # Agrupando o BAI por totalizador de Segmento

    Segto_check = BAI_cascada_R[['Segto_Tot', ANOMES_M0]].groupby(['Segto_Tot']).sum().reset_index()
    Segto_check = Segto_check.merge(BAI_cascada_R, left_on='Segto_Tot', right_on='Sheet', how='left')
    Segto_check['Check_Segto'] = Segto_check[f'{ANOMES_M0}_x'] - Segto_check[f'{ANOMES_M0}_y']
    Segto_check = Segto_check[['Segto_Tot_x', 'Check_Segto']]
    
    

    
    ### 1.3 Total Créditos e Investimentos por Segmento nas Sheets P e P_Ponta
    
    # linhas de Tot Créditos e Invest da Sheet P

    Totais_P = pd.DataFrame()

    for i in range(len(list(segmentos.keys()))):

        df = pmodelo_fin[i].loc[pmodelo_fin[i]['Cod'].isin(['100', '311'])]
        df = df.loc[df['Conceito'].isin(['Comercial'])]
        df['Segto'] = list(segmentos.keys())[i]
        df['Sheet'] = list(segmentos.values())[i]
        Totais_P = Totais_P.append(df,ignore_index = True)

    Totais_P = Totais_P[colunas_r_casc]
    Totais_P['Tipo_Saldo'] = 'P'

    # linhas de Tot Créditos e Invest da Sheet P_Ponta

    Totais_P_Ponta = pd.DataFrame()

    for i in range(len(list(segmentos.keys()))):

        df = pmodelop_fin[i].loc[pmodelop_fin[i]['Cod'].isin(['100', '311'])]
        df = df.loc[df['Conceito'].isin(['Comercial'])]
        df['Segto'] = list(segmentos.keys())[i]
        df['Sheet'] = list(segmentos.values())[i]
        Totais_P_Ponta = Totais_P_Ponta.append(df,ignore_index = True)

    Totais_P_Ponta = Totais_P_Ponta[colunas_r_casc]
    Totais_P_Ponta['Tipo_Saldo'] = 'P_Ponta'

    # Unindo P e P_Ponta

    Totais_P = Totais_P.append(Totais_P_Ponta,ignore_index = True)
    
    # Gerando coluna de totalizadores de Segmento

    Totais_P.loc[Totais_P['Sheet'].isin(segmentos_VarTotal), 'Segto_Tot'] = 'Tot_VarTotal' 
    Totais_P.loc[Totais_P['Sheet'].isin(segmentos_Var), 'Segto_Tot'] = 'Tot_Var' 
    Totais_P.loc[Totais_P['Sheet'].isin(segmentos_PJ), 'Segto_Tot'] = 'Tot_PJ' 
    Totais_P.loc[Totais_P['Sheet'].isin(segmentos_PF), 'Segto_Tot'] = 'Tot_PF'
    
    # Agrupando Total Créditos e Invest de P e P_Ponta por totalizador de Segmento

    Segto_P_check = Totais_P[['Tipo_Saldo', 'Cod', 'Segto_Tot', ANOMES_M0]].groupby(['Tipo_Saldo', 'Cod', 'Segto_Tot']).sum().reset_index()
    Segto_P_check = Segto_P_check.merge(Totais_P, left_on=['Tipo_Saldo', 'Cod', 'Segto_Tot'], right_on=['Tipo_Saldo', 'Cod', 'Sheet'], how='left')
    Segto_P_check['Check_P_Segto'] = Segto_P_check[f'{ANOMES_M0}_x'] - Segto_P_check[f'{ANOMES_M0}_y']
    Segto_P_check = Segto_P_check[['Cod', 'Tipo_Saldo', 'Segto_Tot_x', 'Check_P_Segto']]
    
    
    
    
    ## 2. ANÁLISE
    
    ### 2.1 Comparação MoM de Linhas BPs - deltas acima de |70%| e superiores a |1MM|
    
    # Varejo | Resultado - Linhas BPs (Maiores Variações MoM)
    df = rmodelo_fin[5].loc[rmodelo_fin[5]['Conceito'].isin(conceito_R)]
    df['Segto'] = list(segmentos.keys())[5]
    df['Sheet'] = list(segmentos.values())[5]
    df['MoM Abs'] = df[ANOMES_M0] - df[ANOMES_M1]
    df['MoM %'] = df[ANOMES_M0].div(df[ANOMES_M1].replace(0, 1E-10))
    df = df.loc[(df['MoM %'] >= abs(0.7)) & (df['MoM Abs'] >= abs(1E3))]
    df = df[['Cod', 'Itens / Período', 'Segto', 'Sheet', f'{ANOMES_M2}', f'{ANOMES_M1}', f'{ANOMES_M0}', 'MoM Abs', 'MoM %']]
    df.sort_values(['MoM %'], ascending = False, inplace = True)
    MoM_R01 = df.reset_index()
    
    
    
    
    ### 2.2 Comparação MoM de Linhas BPs - troca de sinal com MoM superiores a |1MM|
    
    # Varejo | Resultado - Linhas BPs (Mudança de Sinal no MoM)

    df = rmodelo_fin[5].loc[rmodelo_fin[5]['Conceito'].isin(conceito_R)]
    df['Segto'] = list(segmentos.keys())[5]
    df['Sheet'] = list(segmentos.values())[5]
    df['MoM Abs'] = df[ANOMES_M0] - df[ANOMES_M1]
    df['MoM %'] = df[ANOMES_M0].div(df[ANOMES_M1].replace(0, 1E-10))
    df = df.loc[df['MoM Abs'] >= abs(1E3)]
    df = df[['Cod', 'Itens / Período', 'Segto', 'Sheet', ANOMES_M1, ANOMES_M0, 'MoM Abs', 'MoM %']]
    df.sort_values(['MoM %'], ascending = False, inplace = True)

    df['BIN_M1'] = df[ANOMES_M1].apply(lambda x: 'Positivo' if x > 0 else 'Negativo' if x < 0 else 'Zero')
    df['BIN_M0'] = df[ANOMES_M0].apply(lambda x: 'Positivo' if x > 0 else 'Negativo' if x < 0 else 'Zero')

    df['check'] = (df['BIN_M1'] == df['BIN_M0'])
    df = df.loc[df['check'] != 'Zero']
    df = df.loc[df['check'] == False]
    MoM_R02 = df.reset_index()
    
    ### 2.3 Linhas BPs com saldo médio negativo
    
    # Saldo Médio - Linhas BPs (Mudança de Sinal no MoM e sinal negativo)

    MoM_P01 = pd.DataFrame()

    for i in range(len(list(segmentos.keys()))):

        df = pmodelo_fin[i].loc[pmodelo_fin[i]['Conceito'].isin(conceito_P)]
        df = df[(df[FY_y0] < 0).any(axis = 1)]
        df['Segto'] = list(segmentos.keys())[i]
        df['Sheet'] = list(segmentos.values())[i]
        df.sort_values([f'{y0}'], ascending = False, inplace = True)
        MoM_P01 = MoM_P01.append(df,ignore_index = True)

    
    return BAI_check, Segto_check, Segto_P_check, MoM_R01, MoM_R02, MoM_P01

def validar_exclusivos_pf_pj(rmod_pf, rmod_pj, cenario):
    m0 = datetime.now().month
    if m0 == 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham'):
        y0 = datetime.now().year - 1
        m0 = 12
        m0 = str(m0)
        
    elif m0 != 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham'):
        y0 = datetime.now().year
        m0 = m0 - 1
        m0 = str(m0)
        if len(m0) == 1:
            m0 = '0' + m0
        
    else:
        y0 = datetime.now().year
        m0 = str(m0)
        if len(m0) == 1:
            m0 = '0' + m0
        
    exclusivo_pj = [
        '54', '57', '10016', '31', '44', '46', '10580', '10561', '10482', '10483', '10484', '10485', '10618', '523', '644', '20160', '921', '970', '1130',
        '41', '38', '394', '530', '37', '388', '529', '34', '35', '52', '10053', '20118', '20120', '20137', '20200', '20117'
    ]

    exclusivo_pf = [
        '22', '40', '17', '18', '19', '10071', '10075', '10588', '10589', '10564', '10565', '10566', '113', '442', '10654', '510', 
        '511', '614', '718', '20109', '20161', '20113', '60933', '701058', '701063', '1087', '1172', '1601185', '12', '13', '14', 
        '10015', '10667', '10668', '10670', '504', '505', '506', '20159', '45','10567', '10308', '390', '537', '1144'
    ]


    rmodelo_pf = deepcopy(rmod_pf)
    rmodelo_pf['Cod'] = rmodelo_pf['Cod'].astype(str)
    rmodelo_pf['Cod'] = rmodelo_pf.apply(lambda row: retirar_ponto_zero(row,'Cod'), axis = 1)
    rmodelo_pf_valid = rmodelo_pf[['Cod', 'Itens / Período', f'{y0}{m0}', f'{y0}']].loc[rmodelo_pf['Cod'].isin(exclusivo_pj)]
    rmodelo_pf_valid = rmodelo_pf_valid.loc[(rmodelo_pf_valid[f'{y0}{m0}'] >= 100) | (rmodelo_pf_valid[f'{y0}'] >= 100)]


    rmodelo_pj = deepcopy(rmod_pj)
    rmodelo_pj['Cod'] = rmodelo_pj['Cod'].astype(str)
    rmodelo_pj['Cod'] = rmodelo_pj.apply(lambda row: retirar_ponto_zero(row,'Cod'), axis = 1)
    rmodelo_pj_valid = rmodelo_pj[['Cod', 'Itens / Período', f'{y0}{m0}', f'{y0}']].loc[rmodelo_pj['Cod'].isin(exclusivo_pf)]
    rmodelo_pj_valid = rmodelo_pj_valid.loc[(rmodelo_pj_valid[f'{y0}{m0}'] >= 100) | (rmodelo_pj_valid[f'{y0}'] >= 100)]
    
    return rmodelo_pf_valid, rmodelo_pj_valid

def validar_contabil_ficticio_cod_zerados_y0(rmodelo_fin, cenario, num_anos):
    m0 = datetime.now().month
    # if m0 == 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham'):
    #     y0 = datetime.now().year - 1
    #     y1 = y0 - 1
    #     y01 = y0 + 1
    #     y02 = y0 + 2
    #     y03 = y0 + 3

    # else:
    y0 = datetime.now().year - 1
    y1 = y0 - 1
    y01 = y0 + 1
    y02 = y0 + 2
    y03 = y0 + 3

    y0 = str(y0)
    y1 = str(y1)
    y01 = str(y01)
    y02 = str(y02)
    y03 = str(y03)

    df1=df2=df3=df4=df5=df6=df7=df8=df9=df10=df11=df12=df13=df14=df15=df16=df17=df18=df19=pd.DataFrame()
    rmodelo_valid_contab_fict = [df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,df14,df15,df16,df17,df18,df19]


    for j in range(len(rmodelo_fin)):
    # Criar lista de Cod únicos que possuem Contábil e Fictício
        cod_contab_fict = rmodelo_fin[j][['Cod']].loc[rmodelo_fin[j]['Conceito'].isin(['Contábil', 'Fictício'])].values
        lst_cod_contab_fict = []
        for i in range(len(cod_contab_fict)):
            if cod_contab_fict[i][-1] != 'x':
                lst_cod_contab_fict.append(cod_contab_fict[i][-1])
        lst_unique_cod_contab_fict = list(set(lst_cod_contab_fict))


        if num_anos == 2:
            all_cols = ['Cod', 'Conceito', 'Itens / Período', f'{y1}', 
                        f'{y0}01', f'{y0}02', f'{y0}03', f'{y0}04', f'{y0}05', f'{y0}06', f'{y0}07', f'{y0}08', f'{y0}09', f'{y0}10', f'{y0}11', f'{y0}12']
        elif num_anos == 3:
            all_cols = ['Cod', 'Conceito', 'Itens / Período', f'{y1}', 
                        f'{y0}01', f'{y0}02', f'{y0}03', f'{y0}04', f'{y0}05', f'{y0}06', f'{y0}07', f'{y0}08', f'{y0}09', f'{y0}10', f'{y0}11', f'{y0}12',
                        f'{y01}01', f'{y01}02', f'{y01}03', f'{y01}04', f'{y01}05', f'{y01}06', f'{y01}07', f'{y01}08', f'{y01}09', f'{y01}10', f'{y01}11', f'{y01}12']
        elif num_anos == 5:
            all_cols = ['Cod', 'Conceito', 'Itens / Período', f'{y1}', 
                        f'{y0}01', f'{y0}02', f'{y0}03', f'{y0}04', f'{y0}05', f'{y0}06', f'{y0}07', f'{y0}08', f'{y0}09', f'{y0}10', f'{y0}11', f'{y0}12',
                        f'{y01}01', f'{y01}02', f'{y01}03', f'{y01}04', f'{y01}05', f'{y01}06', f'{y01}07', f'{y01}08', f'{y01}09', f'{y01}10', f'{y01}11', f'{y01}12',
                        f'{y02}01', f'{y02}02', f'{y02}03', f'{y02}04', f'{y02}05', f'{y02}06', f'{y02}07', f'{y02}08', f'{y02}09', f'{y02}10', f'{y02}11', f'{y02}12',
                        f'{y03}01', f'{y03}02', f'{y03}03', f'{y03}04', f'{y03}05', f'{y03}06', f'{y03}07', f'{y03}08', f'{y03}09', f'{y03}10', f'{y03}11', f'{y03}12']
        else:
            print(f'Número de anos incorreto. Número de anos recebido: {num_anos}, esperava 2, 3 ou 5.')

        # Filtra o rmodelo pelos Cod únicos de contábil e fictício, pelo Conceito contábil e fictício, pelo YTD do ano anterior que é diferente de 0 e não nulo
        df_valid_fin = rmodelo_fin[j][all_cols].loc[(rmodelo_fin[j]['Cod'].isin(lst_unique_cod_contab_fict)) &
                                                (rmodelo_fin[j]['Conceito'].isin(['Contábil', 'Fictício'])) &
                                                (rmodelo_fin[j][f'{y1}'] != 0) &
                                                (pd.isna(rmodelo_fin[j][f'{y1}']) != True)]

        # Dropa as linhas que tem os valores mensais do ano corrente e do ano posterior que são diferentes de 0
        if num_anos == 2:
            df_valid_fin = df_valid_fin.drop(df_valid_fin[(df_valid_fin[f'{y0}01']!=0) &
                                                           (df_valid_fin[f'{y0}02']!=0) &
                                                           (df_valid_fin[f'{y0}03']!=0) &
                                                           (df_valid_fin[f'{y0}04']!=0) &
                                                           (df_valid_fin[f'{y0}05']!=0) &
                                                           (df_valid_fin[f'{y0}06']!=0) &
                                                           (df_valid_fin[f'{y0}07']!=0) &
                                                           (df_valid_fin[f'{y0}08']!=0) &
                                                           (df_valid_fin[f'{y0}09']!=0) &
                                                           (df_valid_fin[f'{y0}10']!=0) &
                                                           (df_valid_fin[f'{y0}11']!=0) &
                                                           (df_valid_fin[f'{y0}12']!=0)].index)       
        
        elif num_anos == 3:
            df_valid_fin = df_valid_fin.drop(df_valid_fin[(df_valid_fin[f'{y0}01']!=0) &
                                                           (df_valid_fin[f'{y0}02']!=0) &
                                                           (df_valid_fin[f'{y0}03']!=0) &
                                                           (df_valid_fin[f'{y0}04']!=0) &
                                                           (df_valid_fin[f'{y0}05']!=0) &
                                                           (df_valid_fin[f'{y0}06']!=0) &
                                                           (df_valid_fin[f'{y0}07']!=0) &
                                                           (df_valid_fin[f'{y0}08']!=0) &
                                                           (df_valid_fin[f'{y0}09']!=0) &
                                                           (df_valid_fin[f'{y0}10']!=0) &
                                                           (df_valid_fin[f'{y0}11']!=0) &
                                                           (df_valid_fin[f'{y0}12']!=0) &
                                                           (df_valid_fin[f'{y01}01']!=0) &
                                                           (df_valid_fin[f'{y01}02']!=0) &
                                                           (df_valid_fin[f'{y01}03']!=0) &
                                                           (df_valid_fin[f'{y01}04']!=0) &
                                                           (df_valid_fin[f'{y01}05']!=0) &
                                                           (df_valid_fin[f'{y01}06']!=0) &
                                                           (df_valid_fin[f'{y01}07']!=0) &
                                                           (df_valid_fin[f'{y01}08']!=0) &
                                                           (df_valid_fin[f'{y01}09']!=0) &
                                                           (df_valid_fin[f'{y01}10']!=0) &
                                                           (df_valid_fin[f'{y01}11']!=0) &
                                                           (df_valid_fin[f'{y01}12']!=0)].index)
        elif num_anos == 5:
            df_valid_fin = df_valid_fin.drop(df_valid_fin[(df_valid_fin[f'{y0}01']!=0) &
                                                           (df_valid_fin[f'{y0}02']!=0) &
                                                           (df_valid_fin[f'{y0}03']!=0) &
                                                           (df_valid_fin[f'{y0}04']!=0) &
                                                           (df_valid_fin[f'{y0}05']!=0) &
                                                           (df_valid_fin[f'{y0}06']!=0) &
                                                           (df_valid_fin[f'{y0}07']!=0) &
                                                           (df_valid_fin[f'{y0}08']!=0) &
                                                           (df_valid_fin[f'{y0}09']!=0) &
                                                           (df_valid_fin[f'{y0}10']!=0) &
                                                           (df_valid_fin[f'{y0}11']!=0) &
                                                           (df_valid_fin[f'{y0}12']!=0) &
                                                           (df_valid_fin[f'{y01}01']!=0) &
                                                           (df_valid_fin[f'{y01}02']!=0) &
                                                           (df_valid_fin[f'{y01}03']!=0) &
                                                           (df_valid_fin[f'{y01}04']!=0) &
                                                           (df_valid_fin[f'{y01}05']!=0) &
                                                           (df_valid_fin[f'{y01}06']!=0) &
                                                           (df_valid_fin[f'{y01}07']!=0) &
                                                           (df_valid_fin[f'{y01}08']!=0) &
                                                           (df_valid_fin[f'{y01}09']!=0) &
                                                           (df_valid_fin[f'{y01}10']!=0) &
                                                           (df_valid_fin[f'{y01}11']!=0) &
                                                           (df_valid_fin[f'{y01}12']!=0) &
                                                           (df_valid_fin[f'{y02}01']!=0) &
                                                           (df_valid_fin[f'{y02}02']!=0) &
                                                           (df_valid_fin[f'{y02}03']!=0) &
                                                           (df_valid_fin[f'{y02}04']!=0) &
                                                           (df_valid_fin[f'{y02}05']!=0) &
                                                           (df_valid_fin[f'{y02}06']!=0) &
                                                           (df_valid_fin[f'{y02}07']!=0) &
                                                           (df_valid_fin[f'{y02}08']!=0) &
                                                           (df_valid_fin[f'{y02}09']!=0) &
                                                           (df_valid_fin[f'{y02}10']!=0) &
                                                           (df_valid_fin[f'{y02}11']!=0) &
                                                           (df_valid_fin[f'{y02}12']!=0) &
                                                           (df_valid_fin[f'{y03}01']!=0) &
                                                           (df_valid_fin[f'{y03}02']!=0) &
                                                           (df_valid_fin[f'{y03}03']!=0) &
                                                           (df_valid_fin[f'{y03}04']!=0) &
                                                           (df_valid_fin[f'{y03}05']!=0) &
                                                           (df_valid_fin[f'{y03}06']!=0) &
                                                           (df_valid_fin[f'{y03}07']!=0) &
                                                           (df_valid_fin[f'{y03}08']!=0) &
                                                           (df_valid_fin[f'{y03}09']!=0) &
                                                           (df_valid_fin[f'{y03}10']!=0) &
                                                           (df_valid_fin[f'{y03}11']!=0) &
                                                           (df_valid_fin[f'{y03}12']!=0)].index)
        else:
            print(f'Número de anos incorreto. Número de anos recebido: {num_anos}, esperava 2, 3 ou 5.')

        # Dropa as colunas que tem soma diferente de 0
        valid_cols = []
        for w in all_cols:
            if w in ('Cod', 'Conceito', 'Itens / Período', f'{y1}') or 0 in (df_valid_fin[w].unique()):
                valid_cols.append(w)
            else:
                pass

        # Seleciona as colunas selecionadas acima para o rmodelo
        rmodelo_valid_contab_fict[j] = df_valid_fin[valid_cols]
        
    return rmodelo_valid_contab_fict

def validar_resultado_e_soma_contabil_ficticio(rmodelo_fin, cenario, num_anos):
    m0 = datetime.now().month
    # if m0 == 1 and cenario.lower() in ('avance', 'avanc', 'avan', 'fechto', 'fechamento', 'fecham'):
    #     y0 = datetime.now().year - 1
    #     y1 = y0 - 1
    #     y01 = y0 + 1
    #     y02 = y0 + 2
    #     y03 = y0 + 3

    # else:
    y0 = datetime.now().year - 1
    y1 = y0 - 1
    y01 = y0 + 1
    y02 = y0 + 2
    y03 = y0 + 3

    y0 = str(y0)
    y1 = str(y1)
    y01 = str(y01)
    y02 = str(y02)
    y03 = str(y03)

    df1=df2=df3=df4=df5=df6=df7=df8=df9=df10=df11=df12=df13=df14=df15=df16=df17=df18=df19=pd.DataFrame()
    rtotal = [df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,df14,df15,df16,df17,df18,df19]
    rcontabil = deepcopy(rtotal)
    rficticio = deepcopy(rtotal)
    rmerge1 = deepcopy(rtotal)
    rmerge2 = deepcopy(rtotal)
    rfinal = deepcopy(rtotal)

    lst1=lst2=lst3=lst4=lst5=lst6=lst7=lst8=lst9=lst10=lst11=lst12=lst13=lst14=lst15=lst16=lst17=lst18=lst19=lst20=[]
    lst_cod_unicos = [lst1,lst2,lst3,lst4,lst5,lst6,lst7,lst8,lst9,lst10,lst11,lst12,lst13,lst14,lst15,lst16,lst17,lst18,lst19,lst20]

    for j in range(len(rmodelo_fin)):
        rcontabil[j] = rmodelo_fin[j].loc[rmodelo_fin[j]['Conceito'].isin(['Contábil'])]
        rficticio[j] = rmodelo_fin[j].loc[rmodelo_fin[j]['Conceito'].isin(['Fictício'])]
        rtotal[j] = rmodelo_fin[j].loc[~rmodelo_fin[j]['Conceito'].isin(['Contábil', 'Fictício'])]

        cod_contab_fict = rmodelo_fin[j][['Cod']].loc[(rmodelo_fin[j]['Conceito'].isin(['Contábil', 'Fictício'])) &
                                                      (~rmodelo_fin[j]['Cod'].isin(['x'])) &
                                                      (pd.isna(rmodelo_fin[j]['Cod']) != True)].values
        lst_cod_contab_fict = []
        for i in range(len(cod_contab_fict)):
            lst_cod_contab_fict.append(cod_contab_fict[i][-1])
        lst_cod_unicos[j] = list(set(lst_cod_contab_fict))

    if num_anos == 2:
        cols = ['Totalizador', 'Chave Cascada', 'Alíquotas', 'Cascada', 'Conceito',
               'Cod', 'Itens / Período', 'Segmentos', 'Linha excel', '', f'Reparto % {y1}',
                f'Reparto % {y0}', 'ISS', 'PIS', 'IR', 'ID',
                f'{y1}01',f'{y1}02',f'{y1}03',f'{y1}04',f'{y1}05',f'{y1}06',f'{y1}07',f'{y1}08',f'{y1}09',f'{y1}10',f'{y1}11',f'{y1}12',f'{y1}',
                f'{y0}01',f'{y0}02',f'{y0}03',f'{y0}04',f'{y0}05',f'{y0}06',f'{y0}07',f'{y0}08',f'{y0}09',f'{y0}10',f'{y0}11',f'{y0}12',f'{y0}']
        
    elif num_anos == 3:
        cols = ['Totalizador', 'Chave Cascada', 'Alíquotas', 'Cascada', 'Conceito',
               'Cod', 'Itens / Período', 'Segmentos', 'Linha excel', '', f'Reparto % {y1}',
                f'Reparto % {y0}', 'ISS', 'PIS', 'IR', 'ID',
                f'{y1}01',f'{y1}02',f'{y1}03',f'{y1}04',f'{y1}05',f'{y1}06',f'{y1}07',f'{y1}08',f'{y1}09',f'{y1}10',f'{y1}11',f'{y1}12',f'{y1}',
               f'{y0}01',f'{y0}02',f'{y0}03',f'{y0}04',f'{y0}05',f'{y0}06',f'{y0}07',f'{y0}08',f'{y0}09',f'{y0}10',f'{y0}11',f'{y0}12',f'{y0}',
               f'{y01}01',f'{y01}02',f'{y01}03',f'{y01}04',f'{y01}05',f'{y01}06',f'{y01}07',f'{y01}08',f'{y01}09',f'{y01}10',f'{y01}11',f'{y01}12',f'{y01}']
        
    elif num_anos == 5:
        cols = ['Totalizador', 'Chave Cascada', 'Alíquotas', 'Cascada', 'Conceito',
               'Cod', 'Itens / Período', 'Segmentos', 'Linha excel', '', f'Reparto % {y1}',
                f'Reparto % {y0}', 'ISS', 'PIS', 'IR', 'ID',
                f'{y1}01',f'{y1}02',f'{y1}03',f'{y1}04',f'{y1}05',f'{y1}06',f'{y1}07',f'{y1}08',f'{y1}09',f'{y1}10',f'{y1}11',f'{y1}12',f'{y1}',
                f'{y0}01',f'{y0}02',f'{y0}03',f'{y0}04',f'{y0}05',f'{y0}06',f'{y0}07',f'{y0}08',f'{y0}09',f'{y0}10',f'{y0}11',f'{y0}12',f'{y0}',
                f'{y01}01',f'{y01}02',f'{y01}03',f'{y01}04',f'{y01}05',f'{y01}06',f'{y01}07',f'{y01}08',f'{y01}09',f'{y01}10',f'{y01}11',f'{y01}12',f'{y01}',
                f'{y02}01', f'{y02}02', f'{y02}03', f'{y02}04', f'{y02}05', f'{y02}06', f'{y02}07', f'{y02}08', f'{y02}09', f'{y02}10', f'{y02}11', f'{y02}12',
                f'{y03}01', f'{y03}02', f'{y03}03', f'{y03}04', f'{y03}05', f'{y03}06', f'{y03}07', f'{y03}08', f'{y03}09', f'{y03}10', f'{y03}11', f'{y03}12']
    else:
        print(f'Número de anos incorreto. Número de anos recebido: {num_anos}, esperava 2, 3 ou 5.')

    for j in range(len(rtotal)):
        rtotal[j] = rtotal[j].loc[rtotal[j]['Cod'].isin(lst_cod_unicos[j])]   
        rcontabil[j] = rcontabil[j].loc[rcontabil[j]['Cod'].isin(lst_cod_unicos[j])]
        rficticio[j] = rficticio[j].loc[rficticio[j]['Cod'].isin(lst_cod_unicos[j])]
        rmerge1[j] = pd.merge(rtotal[j], rcontabil[j], on = 'Cod', how = 'outer')

        for r in cols:
            rmerge1[j] = rmerge1[j].rename(columns={f'{r}_x': f'{r}_result'})
            rmerge1[j] = rmerge1[j].rename(columns={f'{r}_y': f'{r}_contab'})

        rmerge2[j] = pd.merge(rmerge1[j], rficticio[j], on = 'Cod', how = 'outer')

        for p in cols:
            if i != 'Cod':
                rmerge2[j] = rmerge2[j].rename(columns={f'{p}': f'{p}_fict'})

    rfinal = deepcopy(rmerge2)
    for j in range(len(rfinal)):       
        for a in cols[16:]:
            rfinal[j][f'{a}_soma'] = rfinal[j][f'{a}_contab'] + rfinal[j][f'{a}_fict']
        for b in cols:
            if f'{b}_contab' in rfinal[j].columns and f'{b}_fict' in rfinal[j].columns:
                rfinal[j].drop([f'{b}_contab', f'{b}_fict'], axis=1, inplace=True)

        for c in cols[16:]:
            rfinal[j][f'{c}_valid'] = rfinal[j][f'{c}_result'] - rfinal[j][f'{c}_soma']
        for d in cols:
            if f'{d}_result' in rfinal[j].columns and f'{d}_soma' in rfinal[j].columns:
                rfinal[j].drop([f'{d}_result', f'{d}_soma'], axis=1, inplace=True)

        for e in cols:
            rfinal[j] = rfinal[j].rename(columns={f'{e}_result': f'{e}'})
            if e in ('Totalizador','Chave Cascada','Alíquotas','Cascada','Conceito','Segmentos', 'Linha excel', '', f'Reparto % {y1}', f'Reparto % {y0}', 'ISS', 'PIS', 'IR'):
                rfinal[j].drop(e, axis=1, inplace=True)

        for f in cols[16:]:         
            rfinal[j][f'{f}_valid'] = rfinal[j].apply(lambda row: round(row[f'{f}_valid'], 0), axis = 1)  

        # Dropa as linhas que tem os valores mensais do ano corrente e do ano posterior que são diferentes de 0
        if num_anos == 2:
            rfinal[j] = rfinal[j].drop(rfinal[j][((rfinal[j][f'{y1}01_valid']==0) | (pd.isna(rfinal[j][f'{y1}01_valid']) == True)) &
                                                ((rfinal[j][f'{y1}02_valid']==0) | (pd.isna(rfinal[j][f'{y1}02_valid']) == True)) &
                                                ((rfinal[j][f'{y1}03_valid']==0) | (pd.isna(rfinal[j][f'{y1}03_valid']) == True)) &
                                                ((rfinal[j][f'{y1}04_valid']==0) | (pd.isna(rfinal[j][f'{y1}04_valid']) == True)) &
                                                ((rfinal[j][f'{y1}05_valid']==0) | (pd.isna(rfinal[j][f'{y1}05_valid']) == True)) &
                                                ((rfinal[j][f'{y1}06_valid']==0) | (pd.isna(rfinal[j][f'{y1}06_valid']) == True)) &
                                                ((rfinal[j][f'{y1}07_valid']==0) | (pd.isna(rfinal[j][f'{y1}07_valid']) == True)) &
                                                ((rfinal[j][f'{y1}08_valid']==0) | (pd.isna(rfinal[j][f'{y1}08_valid']) == True)) &
                                                ((rfinal[j][f'{y1}09_valid']==0) | (pd.isna(rfinal[j][f'{y1}09_valid']) == True)) &
                                                ((rfinal[j][f'{y1}10_valid']==0) | (pd.isna(rfinal[j][f'{y1}10_valid']) == True)) &
                                                ((rfinal[j][f'{y1}11_valid']==0) | (pd.isna(rfinal[j][f'{y1}11_valid']) == True)) &
                                                ((rfinal[j][f'{y1}12_valid']==0) | (pd.isna(rfinal[j][f'{y1}12_valid']) == True)) &
                                                ((rfinal[j][f'{y0}01_valid']==0) | (pd.isna(rfinal[j][f'{y0}01_valid']) == True)) &
                                                ((rfinal[j][f'{y0}02_valid']==0) | (pd.isna(rfinal[j][f'{y0}02_valid']) == True)) &
                                                ((rfinal[j][f'{y0}03_valid']==0) | (pd.isna(rfinal[j][f'{y0}03_valid']) == True)) &
                                                ((rfinal[j][f'{y0}04_valid']==0) | (pd.isna(rfinal[j][f'{y0}04_valid']) == True)) &
                                                ((rfinal[j][f'{y0}05_valid']==0) | (pd.isna(rfinal[j][f'{y0}05_valid']) == True)) &
                                                ((rfinal[j][f'{y0}06_valid']==0) | (pd.isna(rfinal[j][f'{y0}06_valid']) == True)) &
                                                ((rfinal[j][f'{y0}07_valid']==0) | (pd.isna(rfinal[j][f'{y0}07_valid']) == True)) &
                                                ((rfinal[j][f'{y0}08_valid']==0) | (pd.isna(rfinal[j][f'{y0}08_valid']) == True)) &
                                                ((rfinal[j][f'{y0}09_valid']==0) | (pd.isna(rfinal[j][f'{y0}09_valid']) == True)) &
                                                ((rfinal[j][f'{y0}10_valid']==0) | (pd.isna(rfinal[j][f'{y0}10_valid']) == True)) &
                                                ((rfinal[j][f'{y0}11_valid']==0) | (pd.isna(rfinal[j][f'{y0}11_valid']) == True)) &
                                                ((rfinal[j][f'{y0}12_valid']==0) | (pd.isna(rfinal[j][f'{y0}12_valid']) == True))
                                                ].index)        
        elif num_anos == 3:
            rfinal[j] = rfinal[j].drop(rfinal[j][((rfinal[j][f'{y1}01_valid']==0) | (pd.isna(rfinal[j][f'{y1}01_valid']) == True)) &
                                                ((rfinal[j][f'{y1}02_valid']==0) | (pd.isna(rfinal[j][f'{y1}02_valid']) == True)) &
                                                ((rfinal[j][f'{y1}03_valid']==0) | (pd.isna(rfinal[j][f'{y1}03_valid']) == True)) &
                                                ((rfinal[j][f'{y1}04_valid']==0) | (pd.isna(rfinal[j][f'{y1}04_valid']) == True)) &
                                                ((rfinal[j][f'{y1}05_valid']==0) | (pd.isna(rfinal[j][f'{y1}05_valid']) == True)) &
                                                ((rfinal[j][f'{y1}06_valid']==0) | (pd.isna(rfinal[j][f'{y1}06_valid']) == True)) &
                                                ((rfinal[j][f'{y1}07_valid']==0) | (pd.isna(rfinal[j][f'{y1}07_valid']) == True)) &
                                                ((rfinal[j][f'{y1}08_valid']==0) | (pd.isna(rfinal[j][f'{y1}08_valid']) == True)) &
                                                ((rfinal[j][f'{y1}09_valid']==0) | (pd.isna(rfinal[j][f'{y1}09_valid']) == True)) &
                                                ((rfinal[j][f'{y1}10_valid']==0) | (pd.isna(rfinal[j][f'{y1}10_valid']) == True)) &
                                                ((rfinal[j][f'{y1}11_valid']==0) | (pd.isna(rfinal[j][f'{y1}11_valid']) == True)) &
                                                ((rfinal[j][f'{y1}12_valid']==0) | (pd.isna(rfinal[j][f'{y1}12_valid']) == True)) &
                                                ((rfinal[j][f'{y0}01_valid']==0) | (pd.isna(rfinal[j][f'{y0}01_valid']) == True)) &
                                                ((rfinal[j][f'{y0}02_valid']==0) | (pd.isna(rfinal[j][f'{y0}02_valid']) == True)) &
                                                ((rfinal[j][f'{y0}03_valid']==0) | (pd.isna(rfinal[j][f'{y0}03_valid']) == True)) &
                                                ((rfinal[j][f'{y0}04_valid']==0) | (pd.isna(rfinal[j][f'{y0}04_valid']) == True)) &
                                                ((rfinal[j][f'{y0}05_valid']==0) | (pd.isna(rfinal[j][f'{y0}05_valid']) == True)) &
                                                ((rfinal[j][f'{y0}06_valid']==0) | (pd.isna(rfinal[j][f'{y0}06_valid']) == True)) &
                                                ((rfinal[j][f'{y0}07_valid']==0) | (pd.isna(rfinal[j][f'{y0}07_valid']) == True)) &
                                                ((rfinal[j][f'{y0}08_valid']==0) | (pd.isna(rfinal[j][f'{y0}08_valid']) == True)) &
                                                ((rfinal[j][f'{y0}09_valid']==0) | (pd.isna(rfinal[j][f'{y0}09_valid']) == True)) &
                                                ((rfinal[j][f'{y0}10_valid']==0) | (pd.isna(rfinal[j][f'{y0}10_valid']) == True)) &
                                                ((rfinal[j][f'{y0}11_valid']==0) | (pd.isna(rfinal[j][f'{y0}11_valid']) == True)) &
                                                ((rfinal[j][f'{y0}12_valid']==0) | (pd.isna(rfinal[j][f'{y0}12_valid']) == True)) &
                                                ((rfinal[j][f'{y01}01_valid']==0) | (pd.isna(rfinal[j][f'{y01}01_valid']) == True)) &
                                                ((rfinal[j][f'{y01}02_valid']==0) | (pd.isna(rfinal[j][f'{y01}02_valid']) == True)) &
                                                ((rfinal[j][f'{y01}03_valid']==0) | (pd.isna(rfinal[j][f'{y01}03_valid']) == True)) &
                                                ((rfinal[j][f'{y01}04_valid']==0) | (pd.isna(rfinal[j][f'{y01}04_valid']) == True)) &
                                                ((rfinal[j][f'{y01}05_valid']==0) | (pd.isna(rfinal[j][f'{y01}05_valid']) == True)) &
                                                ((rfinal[j][f'{y01}06_valid']==0) | (pd.isna(rfinal[j][f'{y01}06_valid']) == True)) &
                                                ((rfinal[j][f'{y01}07_valid']==0) | (pd.isna(rfinal[j][f'{y01}07_valid']) == True)) &
                                                ((rfinal[j][f'{y01}08_valid']==0) | (pd.isna(rfinal[j][f'{y01}08_valid']) == True)) &
                                                ((rfinal[j][f'{y01}09_valid']==0) | (pd.isna(rfinal[j][f'{y01}09_valid']) == True)) &
                                                ((rfinal[j][f'{y01}10_valid']==0) | (pd.isna(rfinal[j][f'{y01}10_valid']) == True)) &
                                                ((rfinal[j][f'{y01}11_valid']==0) | (pd.isna(rfinal[j][f'{y01}11_valid']) == True)) &
                                                ((rfinal[j][f'{y01}12_valid']==0) | (pd.isna(rfinal[j][f'{y01}12_valid']) == True))
                                                ].index)
        elif num_anos == 5:
            rfinal[j] = rfinal[j].drop(rfinal[j][((rfinal[j][f'{y1}01_valid']==0) | (pd.isna(rfinal[j][f'{y1}01_valid']) == True)) &
                                                ((rfinal[j][f'{y1}02_valid']==0) | (pd.isna(rfinal[j][f'{y1}02_valid']) == True)) &
                                                ((rfinal[j][f'{y1}03_valid']==0) | (pd.isna(rfinal[j][f'{y1}03_valid']) == True)) &
                                                ((rfinal[j][f'{y1}04_valid']==0) | (pd.isna(rfinal[j][f'{y1}04_valid']) == True)) &
                                                ((rfinal[j][f'{y1}05_valid']==0) | (pd.isna(rfinal[j][f'{y1}05_valid']) == True)) &
                                                ((rfinal[j][f'{y1}06_valid']==0) | (pd.isna(rfinal[j][f'{y1}06_valid']) == True)) &
                                                ((rfinal[j][f'{y1}07_valid']==0) | (pd.isna(rfinal[j][f'{y1}07_valid']) == True)) &
                                                ((rfinal[j][f'{y1}08_valid']==0) | (pd.isna(rfinal[j][f'{y1}08_valid']) == True)) &
                                                ((rfinal[j][f'{y1}09_valid']==0) | (pd.isna(rfinal[j][f'{y1}09_valid']) == True)) &
                                                ((rfinal[j][f'{y1}10_valid']==0) | (pd.isna(rfinal[j][f'{y1}10_valid']) == True)) &
                                                ((rfinal[j][f'{y1}11_valid']==0) | (pd.isna(rfinal[j][f'{y1}11_valid']) == True)) &
                                                ((rfinal[j][f'{y1}12_valid']==0) | (pd.isna(rfinal[j][f'{y1}12_valid']) == True)) &
                                                ((rfinal[j][f'{y0}01_valid']==0) | (pd.isna(rfinal[j][f'{y0}01_valid']) == True)) &
                                                ((rfinal[j][f'{y0}02_valid']==0) | (pd.isna(rfinal[j][f'{y0}02_valid']) == True)) &
                                                ((rfinal[j][f'{y0}03_valid']==0) | (pd.isna(rfinal[j][f'{y0}03_valid']) == True)) &
                                                ((rfinal[j][f'{y0}04_valid']==0) | (pd.isna(rfinal[j][f'{y0}04_valid']) == True)) &
                                                ((rfinal[j][f'{y0}05_valid']==0) | (pd.isna(rfinal[j][f'{y0}05_valid']) == True)) &
                                                ((rfinal[j][f'{y0}06_valid']==0) | (pd.isna(rfinal[j][f'{y0}06_valid']) == True)) &
                                                ((rfinal[j][f'{y0}07_valid']==0) | (pd.isna(rfinal[j][f'{y0}07_valid']) == True)) &
                                                ((rfinal[j][f'{y0}08_valid']==0) | (pd.isna(rfinal[j][f'{y0}08_valid']) == True)) &
                                                ((rfinal[j][f'{y0}09_valid']==0) | (pd.isna(rfinal[j][f'{y0}09_valid']) == True)) &
                                                ((rfinal[j][f'{y0}10_valid']==0) | (pd.isna(rfinal[j][f'{y0}10_valid']) == True)) &
                                                ((rfinal[j][f'{y0}11_valid']==0) | (pd.isna(rfinal[j][f'{y0}11_valid']) == True)) &
                                                ((rfinal[j][f'{y0}12_valid']==0) | (pd.isna(rfinal[j][f'{y0}12_valid']) == True)) &
                                                ((rfinal[j][f'{y01}01_valid']==0) | (pd.isna(rfinal[j][f'{y01}01_valid']) == True)) &
                                                ((rfinal[j][f'{y01}02_valid']==0) | (pd.isna(rfinal[j][f'{y01}02_valid']) == True)) &
                                                ((rfinal[j][f'{y01}03_valid']==0) | (pd.isna(rfinal[j][f'{y01}03_valid']) == True)) &
                                                ((rfinal[j][f'{y01}04_valid']==0) | (pd.isna(rfinal[j][f'{y01}04_valid']) == True)) &
                                                ((rfinal[j][f'{y01}05_valid']==0) | (pd.isna(rfinal[j][f'{y01}05_valid']) == True)) &
                                                ((rfinal[j][f'{y01}06_valid']==0) | (pd.isna(rfinal[j][f'{y01}06_valid']) == True)) &
                                                ((rfinal[j][f'{y01}07_valid']==0) | (pd.isna(rfinal[j][f'{y01}07_valid']) == True)) &
                                                ((rfinal[j][f'{y01}08_valid']==0) | (pd.isna(rfinal[j][f'{y01}08_valid']) == True)) &
                                                ((rfinal[j][f'{y01}09_valid']==0) | (pd.isna(rfinal[j][f'{y01}09_valid']) == True)) &
                                                ((rfinal[j][f'{y01}10_valid']==0) | (pd.isna(rfinal[j][f'{y01}10_valid']) == True)) &
                                                ((rfinal[j][f'{y01}11_valid']==0) | (pd.isna(rfinal[j][f'{y01}11_valid']) == True)) &
                                                ((rfinal[j][f'{y01}12_valid']==0) | (pd.isna(rfinal[j][f'{y01}12_valid']) == True)) &
                                                ((rfinal[j][f'{y02}01_valid']==0) | (pd.isna(rfinal[j][f'{y02}01_valid']) == True)) &
                                                ((rfinal[j][f'{y02}02_valid']==0) | (pd.isna(rfinal[j][f'{y02}02_valid']) == True)) &
                                                ((rfinal[j][f'{y02}03_valid']==0) | (pd.isna(rfinal[j][f'{y02}03_valid']) == True)) &
                                                ((rfinal[j][f'{y02}04_valid']==0) | (pd.isna(rfinal[j][f'{y02}04_valid']) == True)) &
                                                ((rfinal[j][f'{y02}05_valid']==0) | (pd.isna(rfinal[j][f'{y02}05_valid']) == True)) &
                                                ((rfinal[j][f'{y02}06_valid']==0) | (pd.isna(rfinal[j][f'{y02}06_valid']) == True)) &
                                                ((rfinal[j][f'{y02}07_valid']==0) | (pd.isna(rfinal[j][f'{y02}07_valid']) == True)) &
                                                ((rfinal[j][f'{y02}08_valid']==0) | (pd.isna(rfinal[j][f'{y02}08_valid']) == True)) &
                                                ((rfinal[j][f'{y02}09_valid']==0) | (pd.isna(rfinal[j][f'{y02}09_valid']) == True)) &
                                                ((rfinal[j][f'{y02}10_valid']==0) | (pd.isna(rfinal[j][f'{y02}10_valid']) == True)) &
                                                ((rfinal[j][f'{y02}11_valid']==0) | (pd.isna(rfinal[j][f'{y02}11_valid']) == True)) &
                                                ((rfinal[j][f'{y02}12_valid']==0) | (pd.isna(rfinal[j][f'{y02}12_valid']) == True)) &
                                                ((rfinal[j][f'{y03}01_valid']==0) | (pd.isna(rfinal[j][f'{y03}01_valid']) == True)) &
                                                ((rfinal[j][f'{y03}02_valid']==0) | (pd.isna(rfinal[j][f'{y03}02_valid']) == True)) &
                                                ((rfinal[j][f'{y03}03_valid']==0) | (pd.isna(rfinal[j][f'{y03}03_valid']) == True)) &
                                                ((rfinal[j][f'{y03}04_valid']==0) | (pd.isna(rfinal[j][f'{y03}04_valid']) == True)) &
                                                ((rfinal[j][f'{y03}05_valid']==0) | (pd.isna(rfinal[j][f'{y03}05_valid']) == True)) &
                                                ((rfinal[j][f'{y03}06_valid']==0) | (pd.isna(rfinal[j][f'{y03}06_valid']) == True)) &
                                                ((rfinal[j][f'{y03}07_valid']==0) | (pd.isna(rfinal[j][f'{y03}07_valid']) == True)) &
                                                ((rfinal[j][f'{y03}08_valid']==0) | (pd.isna(rfinal[j][f'{y03}08_valid']) == True)) &
                                                ((rfinal[j][f'{y03}09_valid']==0) | (pd.isna(rfinal[j][f'{y03}09_valid']) == True)) &
                                                ((rfinal[j][f'{y03}10_valid']==0) | (pd.isna(rfinal[j][f'{y03}10_valid']) == True)) &
                                                ((rfinal[j][f'{y03}11_valid']==0) | (pd.isna(rfinal[j][f'{y03}11_valid']) == True)) &
                                                ((rfinal[j][f'{y03}12_valid']==0) | (pd.isna(rfinal[j][f'{y03}12_valid']) == True))
                                                ].index)
        else:
            print(f'Número de anos incorreto. Número de anos recebido: {num_anos}, esperava 2, 3 ou 5.')
    return rfinal

def gerar_excel_validadores(valid_tb, valid_contab_fict_zerados_y0, valid_resultado_e_soma_contab_fict, segmentos, cenario, mes_ano_ref, versao):
    path_excel_raw = f"Validadores_Varejo_Atual_{cenario}_{mes_ano_ref}_v{versao}_raw.xlsx"
    with pd.ExcelWriter(path_excel_raw, engine="xlsxwriter") as writer:
        for valid in valid_tb:
            sheet_name = f'{valid=}'.split("=")[0]
            if 'MoM' in sheet_name:
                sheet_name = f'anl_{sheet_name}'
            elif 'rmodelo' in sheet_name:
                sheet_name = sheet_name.replace('rmodelo_', '')
            valid.to_excel(writer, sheet_name=sheet_name, index=False)
        
        cont1 = 0
        for i in valid_resultado_e_soma_contab_fict:
            i.to_excel(writer, sheet_name = f'ContFic_Som_{list(segmentos.values())[cont1]}', index = False)
            cont1 += 1
            
        cont2 = 0
        for i in valid_contab_fict_zerados_y0:
            i.to_excel(writer, sheet_name = f'ContFic_Zer_{list(segmentos.values())[cont2]}', index = False)
            cont2 += 1      
    return path_excel_raw