# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 19:09:02 2021

@author: gusan
"""

from fpdf import FPDF
import easygui
import math

files = easygui.fileopenbox(multiple=True)




H = 297
W = 210

itens = ('ID','N° Ordem','Unidade','Pavimento','Requisição')
dados = ('00','1234','Santana','T','Trocar Janela quebrada')



''' Configuração da pagina '''

pdf= FPDF('P','mm','A4')
pdf.add_page()


''' Linha lateral'''
pdf.set_line_width(5)
pdf.set_draw_color(0,45,0)
pdf.line(10, 10, 10, 290)

''' Header '''
pdf.set_title('TEST')
pdf.set_font('Arial', 'B', 20)
pdf.cell(12)
pdf.cell(50, 10, 'Relatório - Ordem de Serviço')
pdf.ln(15)
pdf.set_line_width(1)
pdf.set_draw_color(0,0,0)
pdf.line(20, 19, 125, 19)

c = 0


'''Informações'''

for inf in itens:
    pdf.set_font('Arial', 'B', 16)
    
    p = pdf.get_y()+10
    pdf.cell(12)
    pdf.cell(50, p, f'{inf}: ')
    es = pdf.get_x()+3
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(es, p, dados[c])
    pdf.ln(10)
    c+=1

p = pdf.get_y()+10
pdf.cell(12)
pdf.set_font('Arial', 'B', 16)
pdf.cell(50, p, 'Observações')

#linha observação
p = pdf.get_y()+50
es = pdf.get_x()+1
pdf.set_line_width(1)
pdf.set_draw_color(0,0,0)
pdf.line(20, p, 125, p)

pdf.ln(1)

pdf.cell(12)
es = pdf.get_x()+1
p = pdf.get_y()+40
pdf.set_font('Arial', 'B', 12)

pdf.multi_cell(50, 210, 'Rachadura na contra verga -- sobre peso na laje -- novo pilar necessario',1)



'''Imagens'''


pdf.add_page()
# pdf.ln(10)
# pdf.cell(12)
# p = pdf.get_y()+40
pdf.set_font('Arial', 'B', 16)
pdf.cell(50, 10, 'Imagens')

x= 1
cont = 1
p = 30

test = len(files)

play = math.ceil(test/2)

for img in files:
    # p = pdf.get_y()+80
    verificador = x%play
    
    if verificador == 0:
        pdf.image(img,10,p,w=180,h=90)
        pdf.ln(1)
        x+=1
        p+=100
        if  cont != test:
            print ('Break\n')
            pdf.add_page()
            p = pdf.get_y()+30
            cont+=1
    else:    
        pdf.image(img,10,p,w=180,h=90)
        pdf.ln(1)
        x+=1
        p+=100
        cont+=1
        # pdf.accept_page_break()
        



pdf.output('tuto2.pdf', 'F')