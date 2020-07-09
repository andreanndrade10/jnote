import seaborn as sns
import matplotlib.pyplot as plt
from docx import Document
from jnote import *

def write_table_rede_cabeada(school_name):
    index = df_questions[df_questions['Unidade']==school_name].index.tolist()[0]

    records = (
        ('Fabricante',df_questions[14][index],df_questions[23][index],df_questions[32][index],df_questions[43][index]),
        ('Modelo',df_questions[15][index],df_questions[24][index],df_questions[33][index],df_questions[44][index]),
        ('Total de portas',df_questions[19][index],df_questions[28][index],df_questions[37][index],'---'),
        ('Total de SW do modelo',df_questions[17][index],df_questions[26][index],df_questions[35][index],df_questions[45][index])
    )

    table = document.add_table(rows=1, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '        '
    hdr_cells[1].text = 'Modelo 3'
    hdr_cells[2].text = 'Modelo 2'
    hdr_cells[3].text = 'Modelo 1'
    hdr_cells[4].text = 'Switch L3'

    for a,b,c,d,e in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(a)
        row_cells[1].text = str(b)
        row_cells[2].text = str(c)
        row_cells[3].text = str(d)
        row_cells[4].text = str(e)

def write_table_wireless(school_name):
    index = df_questions[df_questions['Unidade']==school_name].index.tolist()[0]
    records = (
        ('Fabricante',df_questions[64][index], df_questions[71][index]),
        ('Modelo', df_questions[65][index], df_questions[72][index]),
        ('MU-MIMO', df_questions[70][index], '---'),
        ('Quantidade de APs Suportados','---',df_questions[73][index]),
        ('Quantidade de Clientes total suportada','---',df_questions[74][index])
    )

    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '        '
    hdr_cells[1].text = 'Access Point'
    hdr_cells[2].text = 'Controladora'

    for a,b,c in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(a)
        row_cells[1].text = str(b)
        row_cells[2].text = str(c)
 
def write_table_conectividade(school_name):
    index = df_questions[df_questions['Unidade']==school_name].index.tolist()[0]
    records = (
        ('Banda do Link',df_questions[79][index], df_questions[84][index]),
        ('Tipo de Link', df_questions[78][index], df_questions[83][index]),
        ('Provedor', df_questions[80][index], df_questions[85][index])
    )

    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '        '
    hdr_cells[1].text = 'Link Primário (Mbps)'
    hdr_cells[2].text = 'Link Secundário (Mbps)'

    for a,b,c in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(a)
        row_cells[1].text = str(b)
        row_cells[2].text = str(c)

def write_table_controle_de_conteudo(school_name):
    index = df_questions[df_questions['Unidade']==school_name].index.tolist()[0]
    records = (
        ('Modelo', df_questions[90][index], df_questions[97][index]),
        ('Fabricante', df_questions[89][index], '---'),
        ('Tipo', df_questions[88][index], '---')
    )

    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '        '
    hdr_cells[1].text = 'Firewall'
    hdr_cells[2].text = 'NAC'

    for a,b,c in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(a)
        row_cells[1].text = str(b)
        row_cells[2].text = str(c)

def write_unit_tables(school_name):
    document.add_paragraph('Rede Cabeada', style='List Continue 3')
    write_table_rede_cabeada(school_name)
    document.add_paragraph('Wireless', style='List Continue 3')
    write_table_wireless(school_name)
    document.add_paragraph('Conectividade', style='List Continue 3')
    write_table_conectividade(school_name)
    document.add_paragraph('Controle de Conteúdo', style='List Continue 3')
    write_table_controle_de_conteudo(school_name)

def write_specifics(school_name):
    document.add_paragraph(specific_infra(school_name), style='List Continue 2')
    document.add_paragraph(specific_connection(school_name), style='List Continue 2')
    document.add_paragraph(specific_contentMaturity(school_name), style='List Continue 2')
    #document.add_paragraph('Especificos', style = 'List Number')


def write_percentage_class(lugar):
    explanation = 'Distribuição da classificação das escolas em {}'.format(lugar)
    document.add_paragraph(explanation)
    classe = ['A','B','C','D','E','F','G','H','I']
    for i in range(len(class_percentage(lugar))):
        if class_percentage(lugar)[i] != 0:
            text = 'Classe ' + classe[i] + ': ' + str(class_percentage(lugar)[i]) + '%'
            document.add_paragraph(text, style='List 2')

def write_percentage_size(lugar):
    explanation = 'Distribuição de tamanho das esolas em {}'.format(lugar)
    document.add_paragraph(explanation)
    size = ['Grande', 'Média', 'Pequena']
    for i in range(len(size_percentage(lugar))):
        if size_percentage(lugar)[i] != 0:
            text =  size[i] + ': ' + str(size_percentage(lugar)[i]) + '%'
            document.add_paragraph(text, style='List 2')


def write_regiao(regiao):
    document.add_heading(regiao, level=1)
    write_percentage_class(regiao)
    write_percentage_size(regiao)

def write_estado(estado):
    document.add_heading(estado, level=2)
    write_percentage_class(estado)
    write_percentage_size(estado)

def write_unidades(unidade):
    document.add_paragraph(unidade, style='List Bullet')
    write_unit_tables(unidade)
    a = document.add_paragraph('Classe pertencente: ', style='List 2')
    a.add_run(get_classe(unidade)) 
    p = document.add_paragraph('Maior prioridade de mudança: ', style='List 2')
    p.add_run(change_priority(unidade)).italic = True
    write_specifics(unidade)



def write_doc():
    for i in df['Regiao'].unique():
        write_regiao(i)
        for j in df[df.Regiao==i]['UF'].unique():
            write_estado(j)
            for k in df[df.UF==j]['Unidade']:
                write_unidades(k)
                


if __name__ == "__main__":
    document = Document()
    write_doc()
    document.save('escolas.docx')