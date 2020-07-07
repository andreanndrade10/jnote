import pandas as pd
import numpy as np
from decimal import Decimal

df_region = pd.read_excel('db.xlsx', 'regiao')
df_region['Regiao'] = df_region['Regiao'].map(lambda regiao: regiao.split()[-1])
df = pd.read_excel('db.xlsx','classificacao')
df = pd.merge(df, df_region[['UF','Regiao']],how='inner' ,on='UF')
df_questions = pd.read_excel('db.xlsx', 'questions')

# input: Brasil, Norte, Nordeste, Sul, BA, DF, SP...
# return percentage list classA, classB, classC, ... 
def class_percentage(lugar):
    if lugar.capitalize() in ['Norte', 'Sul', 'Nordeste', 'Centro-oeste', 'Sudeste']:
        return class_percentage_regiao(lugar)
    elif lugar.upper() in ['AC','AP','AM','BA','CE','DF','ES','GO',
                           'MA','MT','MS','MG','PA','PB','PR','PE',
                           'PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO']:
        return class_percentage_estado(lugar)
    else:
        return class_percentage_general()
 

def class_percentage_general():
    percent_class_A_to_I = []
    classes = ['A','B','C','D','E','F','G','H','I']

    for i in classes:
        percent_class_A_to_I.append(round((len(df[df.Classe == 'Classe {}'.format(i)]) / len(df))*100, 2))
    return percent_class_A_to_I

'''
def class_percentage_general():
    percent_class_A = Decimal((len(df[df.Classe == 'Classe A']) / len(df))*100)
    percent_class_B = Decimal((len(df[df.Classe == 'Classe B']) / len(df))*100)
    percent_class_C = Decimal((len(df[df.Classe == 'Classe C']) / len(df))*100)
    percent_class_D = Decimal((len(df[df.Classe == 'Classe D']) / len(df))*100)
    percent_class_E = Decimal((len(df[df.Classe == 'Classe E']) / len(df))*100)
    percent_class_F = Decimal((len(df[df.Classe == 'Classe F']) / len(df))*100)
    percent_class_G = Decimal((len(df[df.Classe == 'Classe G']) / len(df))*100)
    percent_class_H = Decimal((len(df[df.Classe == 'Classe H']) / len(df))*100)
    percent_class_I = Decimal((len(df[df.Classe == 'Classe I']) / len(df))*100)
    
    return round(percent_class_A,2), round(percent_class_B,2),round(percent_class_C,2), round(percent_class_D,2), round(percent_class_E,2),round(percent_class_F,2),round(percent_class_G,2), round(percent_class_H,2), round(percent_class_I,2)
'''

def class_percentage_regiao(lugar):
    regiao_group = df[df.Regiao == lugar]
    percent_class_A_to_I = []
    classes = ['A','B','C','D','E','F','G','H','I']

    for i in classes:
        percent_class_A_to_I.append(round((len(regiao_group[regiao_group.Classe=='Classe {}'.format(i)]) / (regiao_group['Classe'].count()))*100, 2))
    return percent_class_A_to_I

'''
def class_percentage_regiao(lugar):
    regiao_group = df[df.Regiao == lugar]
    percent_class_A = (len(regiao_group[regiao_group.Classe=='Classe A']) / (regiao_group['Classe'].count()))*100
    percent_class_B = (len(regiao_group[regiao_group.Classe=='Classe B']) / (regiao_group['Classe'].count()))*100
    percent_class_C = (len(regiao_group[regiao_group.Classe=='Classe C']) / (regiao_group['Classe'].count()))*100
    percent_class_D = (len(regiao_group[regiao_group.Classe=='Classe D']) / (regiao_group['Classe'].count()))*100
    percent_class_E = (len(regiao_group[regiao_group.Classe=='Classe E']) / (regiao_group['Classe'].count()))*100
    percent_class_F = (len(regiao_group[regiao_group.Classe=='Classe F']) / (regiao_group['Classe'].count()))*100
    percent_class_G = (len(regiao_group[regiao_group.Classe=='Classe G']) / (regiao_group['Classe'].count()))*100
    percent_class_H = (len(regiao_group[regiao_group.Classe=='Classe H']) / (regiao_group['Classe'].count()))*100
    percent_class_I = (len(regiao_group[regiao_group.Classe=='Classe I']) / (regiao_group['Classe'].count()))*100
    return round(percent_class_A,2), round(percent_class_B,2),round(percent_class_C,2), round(percent_class_D,2), round(percent_class_E,2),round(percent_class_F,2),round(percent_class_G,2), round(percent_class_H,2), round(percent_class_I,2)
'''
def class_percentage_estado_2(lugar):
    estado_group = df[df.UF == lugar]
    percent_class_A_to_I = []
    classes = ['A','B','C','D','E','F','G','H','I']

    for i in classes:
        percent_class_A_to_I.append(round((len(estado_group[estado_group.Classe=='Classe {}'.format(i)]) / (estado_group['Classe'].count()))*100 ,2))
    return percent_class_A_to_I

def class_percentage_estado(lugar):
    estado_group = df[df.UF == lugar]
    percent_class_A = (len(estado_group[estado_group.Classe=='Classe A']) / (estado_group['Classe'].count()))*100
    percent_class_B = (len(estado_group[estado_group.Classe=='Classe B']) / (estado_group['Classe'].count()))*100
    percent_class_C = (len(estado_group[estado_group.Classe=='Classe C']) / (estado_group['Classe'].count()))*100
    percent_class_D = (len(estado_group[estado_group.Classe=='Classe D']) / (estado_group['Classe'].count()))*100
    percent_class_E = (len(estado_group[estado_group.Classe=='Classe E']) / (estado_group['Classe'].count()))*100
    percent_class_F = (len(estado_group[estado_group.Classe=='Classe F']) / (estado_group['Classe'].count()))*100
    percent_class_G = (len(estado_group[estado_group.Classe=='Classe G']) / (estado_group['Classe'].count()))*100
    percent_class_H = (len(estado_group[estado_group.Classe=='Classe H']) / (estado_group['Classe'].count()))*100
    percent_class_I = (len(estado_group[estado_group.Classe=='Classe I']) / (estado_group['Classe'].count()))*100
    return round(percent_class_A,2), round(percent_class_B,2),round(percent_class_C,2), round(percent_class_D,2), round(percent_class_E,2),round(percent_class_F,2),round(percent_class_G,2), round(percent_class_H,2), round(percent_class_I,2)

# Return list with percentage (grande, média, pequena)
def size_percentage(lugar):
    if lugar.capitalize() in ['Norte', 'Sul', 'Nordeste', 'Centro-oeste', 'Sudeste']:
        return size_percentage_regiao(lugar)
    elif lugar.upper() in ['AC','AP','AM','BA','CE','DF','ES','GO',
                           'MA','MT','MS','MG','PA','PB','PR','PE',
                           'PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO']:
        return size_percentage_estado(lugar)
    else:
        return size_percentage_general()

def size_percentage_general():
    percentage_pequena = (len(df[df.Tamanho == 'Pequena']) / len(df))*100
    percentage_media = (len(df[df.Tamanho == 'Média']) / len(df))*100
    percentage_grande = (len(df[df.Tamanho == 'Grande']) / len(df))*100
    return round(percentage_grande,2), round(percentage_media,2), round(percentage_pequena,2)
    
def size_percentage_regiao(lugar):
    regiao_group = df[df.Regiao==lugar]
    percentage_pequena = (len(regiao_group[regiao_group.Tamanho == 'Pequena'])/len(regiao_group))*100
    percentage_media = (len(regiao_group[regiao_group.Tamanho == 'Média'])/len(regiao_group))*100
    percentage_grande = (len(regiao_group[regiao_group.Tamanho == 'Grande'])/len(regiao_group))*100
    return round(percentage_grande,2), round(percentage_media,2), round(percentage_pequena,2)

def size_percentage_estado(lugar):
    estado_group = df[df.UF==lugar]
    percentage_pequena = (len(estado_group[estado_group.Tamanho == 'Pequena'])/len(estado_group))*100
    percentage_media = (len(estado_group[estado_group.Tamanho == 'Média'])/len(estado_group))*100
    percentage_grande = (len(estado_group[estado_group.Tamanho == 'Grande'])/len(estado_group))*100
    return round(percentage_grande,2), round(percentage_media,2), round(percentage_pequena,2)
    
# Escolaas Individualmente
def change_priority(school_name):
    maturities = []
    #pandas_index = df.index[df['Unidade'] == 'CFP Candeias'].tolist()
    grades = ['Maturidade Infra', 'Maturidade Conexão', 'Maturidade Controle de Conteúdo']
    index = df[df['Unidade']==school_name].index.tolist()[0]

    for i in grades:
        maturities.append(df[i][index])
    
    index_smaller = maturities.index(min(maturities))     
    #return index_smaller
    #index smaller points the most fragile aspect of the network (0:infra , 1:connection, 2: content control)
    
    if index_smaller == 0:
        return 'infraestrutura'
    elif index_smaller ==1:
        return 'conexão'
    elif index_smaller ==2:
        return 'controle de conteúdo'
    else:
        print('Something wrong in change_priority() function')

def get_classe(name_school):
    index = df[df['Unidade']==name_school].index.tolist()[0]
    classe = df['Classe'][index]
    return classe

def specific_infra(name_school):
    questions = [36,38,47,68]
    index = df_questions[df_questions['Unidade']==name_school].index.tolist()[0]
    evaluation = []
    
    for i in questions:
        evaluation.append(df_questions[i][index])
        
    if evaluation[0] == 'Não':
        return 'Switches não possuem portas Gigabit Ethernet'
    if evaluation[1] == 'Não':
        return 'Switches não suportam PoE'
    if evaluation[2] == 'Não':
        return 'Switches não suportam roteamento entre VLANs'
    if '802.11ac' not in evaluation[3].split(';'):
        return 'Protocolos suportados pelos APs podem ser melhorados'

def specific_connection(name_school):
    questions = [78,79,84]
    index = df_questions[df_questions['Unidade']==name_school].index.tolist()[0]
    evaluation = []
    
    for i in questions:
        evaluation.append(df_questions[i][index])
    
    if evaluation[0] != 'MPLS' and evaluation[0] != 'Dedicado':
        msg1 = 'Link não recomendado, pode ser substituido por MPLS ou por link dedicado.'
    else:
        msg1 = ''
    # Determinar com o Laurenz o que é uma velocidade boa;
    if not check_link_quality(name_school, evaluation[1]): 
        msg2 = ' Velocidade da banda contratada abaixo do recomendado.'
    else:
        msg2 = ''
    return msg1+msg2 

def check_link_quality(name_school, bandwidth):
    index_question = df_questions[df_questions['Unidade']==name_school].index.tolist()[0]
    index_df = df[df['Unidade']==name_school].index.tolist()[0]
    school_size = df['Tamanho'][index_df]

    if school_size == 'Grande' and df_questions[79][index_question]<100:
        return False
    elif school_size == 'Média' and df_questions[79][index_question]<70:
        return False
    elif school_size == 'Pequena' and df_questions[79][index_question]<40:
        return False
    else:
        return True

# In this case, we will use the DataFrame 'df' not df_questions!!!
def specific_contentMaturity(name_school):
    index = df[df['Unidade']==name_school].index.tolist()[0]
    grade_content_maturity = df['Maturidade Controle de Conteúdo'][index]
    
    if grade_content_maturity == 0:
        return 'Infraestrutura carece de Firewall e NAC'
    if grade_content_maturity == 3:
        return 'Infraestrutura possui NAC, mas carece de Firewall'
    if grade_content_maturity == 4:
        return 'Infraestrutura possui firewall, mas carece de Proxy ou NAC'
    if grade_content_maturity == 7:
        return 'Infraestrutura possui Next Generation Firewall sem NAC ou Firewall/Proxy com NAC'