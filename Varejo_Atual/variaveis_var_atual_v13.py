# Executar os comandos no Terminal, um de cada vez (sem '#')
#conda remove openpyxl
#conda install openpyxl==3.0.1

# Pivot Table
pivot_table_file = 'TB_ANALITICO_CENARIO_PNL.csv'
# pivot_table_file = 'TB_ANALITICO_CENARIO_PNL_AVANCE.xlsx'


# RModelo
r_modelo_file = 'Modelo_BP_Atual.xlsx'
r_sheet_name = 'R_MODELO'
r_skip_rows = 9
r_nrows = 10019
# r_nrows = 2918
r_use_cols = "A:AV"


# PModelo
p_modelo_file = 'Modelo_BP_Atual.xlsx'
p_sheet_name = 'P_MODELO'
p_skip_rows = 9
p_nrows = 3669
p_use_cols = "A:AW"


# SModelo
s_modelo_file = 'Modelo_BP_Atual.xlsx'
s_sheet_name = 'S_MODELO'
s_skip_rows = 9
s_nrows = 969
#s_use_cols = "F:AP"
s_use_cols = "A:AW"
s_anos = '3'  # Número de anos de spread para ser calculado. Caso a base pivot_table_file tenha mais anos e você selecionar menos anos para spread, o restante dos anos virá zerado.
## PARA O P27 MUDAR PARA 5 POIS SE CALCULA 5 ANOS 


# CascModelo
c_modelo_file = 'Modelo_BP_Atual.xlsx'
c_sheet_name = 'Casc_MODELO'
c_skip_rows = 9
c_nrows = 128
c_use_cols = "F:G"


# Validadores (somente para cenario P)
anomes_m0 = '202412'
anomes_m1 = '202411'
anomes_m2 = '202410'


# Excel Final
cenario = 'Previa'  # ('Previa', 'P28', 'Avance', 'Fechto')
mesano = 'Out'  ## para o p27 Jul24
versao = '1'      ## ultimo P28: 12


# Variáveis Gerais 28 segmentos antes(2024) tinha 19 segmentos
segmentos = {
 ##   'Varejo Total':'Tot_VarTotal',        ## nao tem na lista da tabela do oracle
    'Real Estate.':'Tot_RealEstate',      ## OK - 2024 e 2025
   ## 'Toro Canal Externo':'Tot_Toro',    ## OK - 2024 e 2025                 -- REMOVER
    'Gira':'Tot_Gira',                    ## OK - 2024 e 2025
    'Varejo':'Tot_Var',                   ## OK - 2024 e 2025   
    'PF':'Tot_PF',                        ## OK - 2024 e 2025 #### TOTAL PF   --ADICIONADO 1
    'PF Ex - Agro':'Tot_PF_Ex_Agro',      ##novo 14/10/2025
    'Especial Total': 'Tot_Esp',          ## 2025   #########  totalizador 
    'Especial': 'Esp',                    ## 2025   #########  PF             --ADICIONADO 1
    'Prospera': 'Prosp',                  ## 2025   #########  PF 
   ## 'PF1+PF2':'Tot_PF2_PF1',            ## OK - 2024 e 2025  totalizador    -- REMOVER
   ## 'PF I':'PF1',                       ## OK - 2024 e 2025  PF             -- REMOVER 1
   ## 'PF II':'PF2',                      ## OK - 2024 e 2025  PF             -- REMOVER 1
    'Select Total': 'Tot_Sel',            ## 2025   #########  totalizador
    'Select Ex - Agro': 'Tot_Sel_ExAgro',            ## 2025   #########  totalizador NEW 13-03-2025
    'Agro Select': 'Agro',                ## 2025   #########  PF
    'Select':'Select',                    ## OK - 2024 e 2025  PF
    'Select High':'Select_Hi',            ## OK - 2024 e 2025  PF
    'PJ':'Tot_PJ',                        ## OK - 2024 e 2025 #### TOTAL PJ
    'Empresas':'Tot_Emp',                 ## OK - 2024 e 2025  totalizador
    'Empresas I':'Emp1',                  ## OK - 2024 e 2025  PJ
    'Empresas II':'Emp2',                 ## OK - 2024 e 2025  PJ
    'Empresas III':'Emp3',                ## OK - 2024 e 2025  PJ
    'Empresas Digital':'Tot_Emp_Dig',     ## 2025   #########  totalizador
    'Empresas I Massivo':'Emp1_Mass',     ## 2025   #########  PJ
    'Empresas MEI':'Emp_MEI',             ## OK - 2024 e 2025  PJ
    'GIU':'Tot_Giu',                      ## 2025   #########  PJ
    'Governos':'Tot_Gov',                 ## OK - 2024 e 2025  PJ obs: duvida se mantem ou nao
    'Universidades':'Tot_Univ',           ## OK - 2024 e 2025  PJ obs: duvida se mantem ou nao
}

# segmentos = {
#     'Varejo Ampliado':'Tot_VarTotal',
#     'Consumer Finance':'Tot_Financ_Total',
#     'Real Estate.':'Tot_RealEstate',
#     'Toro Canal Externo':'Tot_Toro',
#     'Gira':'Tot_Gira',
#     'Varejo':'Tot_Var',
#     'PF1+PF2':'Tot_PF2_PF1',
#     'PF':'Tot_PF',
#     'PF I':'PF1',
#     'PF II':'PF2',
#     'Van Gogh':'VG',
#     'Select':'Select',
#     'PJ':'Tot_PJ',
#     'Empresas':'Tot_Emp',
#     'Empresas I':'Emp1',
#     'Empresas II':'Emp2',
#     'Empresas III':'Emp3',
#     'Empresas MEI':'Emp_MEI',
#     'Governos':'Tot_Gov',
#     'Universidades':'Tot_Univ',
# }

#subtipos = ['Comercial', 'Morosa Total', 'Carteira total', 'Outros', 'Acordos', 'Resultado Total', 'Resultado contabil', 'TTI Total']
subtipos = ['Comercial', 'Morosa Total', 'Carteira total', 'Outros', 'Acordos', 'Resultado Total', 'Contábil', 'Fictício']

dic_dias_mes = {
    '202201':31,
    '202202':28,
    '202203':31,
    '202204':30,
    '202205':31,
    '202206':30,
    '202207':31,
    '202208':31,
    '202209':30,
    '202210':31,
    '202211':30,
    '202212':31,
    '2022'  :365,
    
    '202301':31,
    '202302':28,
    '202303':31,
    '202304':30,
    '202305':31,
    '202306':30,
    '202307':31,
    '202308':31,
    '202309':30,
    '202310':31,
    '202311':30,
    '202312':31,
    '2023'  :365,
    
    '202401':31,
    '202402':29,
    '202403':31,
    '202404':30,
    '202405':31,
    '202406':30,
    '202407':31,
    '202408':31,
    '202409':30,
    '202410':31,
    '202411':30,
    '202412':31,
    '2024'  :365,
    
    '202501':31,
    '202502':28,
    '202503':31,
    '202504':30,
    '202505':31,
    '202506':30,
    '202507':31,
    '202508':31,
    '202509':30,
    '202510':31,
    '202511':30,
    '202512':31,
    '2025'  :365,
    
    '202601':31,
    '202602':28,
    '202603':31,
    '202604':30,
    '202605':31,
    '202606':30,
    '202607':31,
    '202608':31,
    '202609':30,
    '202610':31,
    '202611':30,
    '202612':31,
    '2026'  :365,
    
    '202701':31,
    '202702':28,
    '202703':31,
    '202704':30,
    '202705':31,
    '202706':30,
    '202707':31,
    '202708':31,
    '202709':30,
    '202710':31,
    '202711':30,
    '202712':31,
    '2027'  :365,
    
    '202801':31,
    '202802':28,
    '202803':31,
    '202804':30,
    '202805':31,
    '202806':30,
    '202807':31,
    '202808':31,
    '202809':30,
    '202810':31,
    '202811':30,
    '202812':31,
    '2028'  :365

}

dic_du_mes = {
    '202201':21,
    '202202':19,
    '202203':22,
    '202204':19,
    '202205':22,
    '202206':21,
    '202207':21,
    '202208':23,
    '202209':21,
    '202210':20,
    '202211':20,
    '202212':22,
    '2022'  :252,
    
    '202301':22,
    '202302':18,
    '202303':23,
    '202304':18,
    '202305':22,
    '202306':21,
    '202307':21,
    '202308':23,
    '202309':20,
    '202310':21,
    '202311':20,
    '202312':20,
    '2023'  :252,
    
    '202401':22,
    '202402':19,
    '202403':20,
    '202404':22,
    '202405':21,
    '202406':20,
    '202407':23,
    '202408':22,
    '202409':21,
    '202410':23,
    '202411':20,
    '202412':21,
    '2024'  :252,
    
    '202501':22,
    '202502':20,
    '202503':19,
    '202504':20,
    '202505':21,
    '202506':20,
    '202507':23,
    '202508':21,
    '202509':22,
    '202510':23,
    '202511':19,
    '202512':22,
    '2025'  :252,
    
    '202601':21,
    '202602':18,
    '202603':22,
    '202604':20,
    '202605':20,
    '202606':21,
    '202607':23,
    '202608':21,
    '202609':21,
    '202610':21,
    '202611':19,
    '202612':22,
    '2026'  :252,
    
    '202701':20,
    '202702':18,
    '202703':22,
    '202704':21,
    '202705':20,
    '202706':22,
    '202707':22,
    '202708':22,
    '202709':21,
    '202710':20,
    '202711':20,
    '202712':23,
    '2027'  :252,
    
    '202801':21,
    '202802':19,
    '202803':23,
    '202804':18,
    '202805':22,
    '202806':21,
    '202807':21,
    '202808':23,
    '202809':20,
    '202810':21,
    '202811':19,
    '202812':20,
    '2028'  :252
}