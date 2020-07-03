from docx import Document
from jnote import *

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