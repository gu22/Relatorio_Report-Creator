# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 19:09:02 2021

@author: gusan
"""

from fpdf import FPDF
import easygui

files = easygui.fileopenbox(multiple=True)

H = 297
W = 210

itens = ('ID','N° Ordem','Unidade','Pavimento','Requisição')
dados = ('00','1234','Santana','T','Trocar Janela quebrada')


pdf= FPDF('P','mm','A4')
pdf.add_page()


#Linha lateral
pdf.set_line_width(5)
pdf.set_draw_color(0,45,0)
pdf.line(10, 10, 10, 290)

#Header
pdf.set_title('TEST')
pdf.set_font('Arial', 'B', 20)
pdf.cell(12)
pdf.cell(50, 10, 'Relatório - Ordem de Serviço')
pdf.ln(15)
pdf.set_line_width(1)
pdf.set_draw_color(0,0,0)
pdf.line(20, 19, 125, 19)

c = 0
#Informações
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
pdf.cell(50, p, 'Rachadura na contra verga -- sobre peso na laje -- novo pilar necessario')


#Imagens
pdf.add_page()
# pdf.ln(10)
# pdf.cell(12)
# p = pdf.get_y()+40
pdf.set_font('Arial', 'B', 16)
pdf.cell(50, p, 'Imagens')

x= 0
for img in files:
    # p = pdf.get_y()+80
    if x == 0:
        p = pdf.get_y()+80
        pdf.image(img,10,p,w=180,h=90)
        pdf.ln(1)
        x+=1
    else:
        pdf.image(img,10,(p+80),w=180,h=90)
        




pdf.output('tuto2.pdf', 'F')