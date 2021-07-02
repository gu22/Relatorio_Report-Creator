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

itens = ('ID','N° Ordem','Unidade','Pavimento','Requisição','')
dados = ('00','1234','Santana','T','Trocar Janela quebrada','')



''' Configuração da pagina '''

pdf= FPDF('P','mm','A4')
# pdf.set_right_margin(10)
# pdf.page_no()
pdf.alias_nb_pages()
pdf.set_auto_page_break(True,10)

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

''' Footer'''

def footer(self):
    # Position at 1.5 cm from bottom
        self.set_y(-15)
        # helvetica italic 8
        self.set_font('helvetica', 'I', 8)
        # Page number
        self.cell(0, 0, 'Página ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

pdf.cell(0,)

'''Informações'''

for inf in itens:
    pdf.set_font('Arial', 'B', 16)
    
    p = pdf.get_y()+10
    pdf.cell(12)
    pdf.cell(50, p, f'{inf}: ')
    es = pdf.get_x()+3
    pdf.set_font('Arial', 'I', 12)
    pdf.cell(es, p, dados[c])
    pdf.ln(10)
    c+=1
    
pdf.ln(55)
p = pdf.get_y()+10
pdf.cell(12)
pdf.set_font('Arial', 'B', 16)
pdf.cell(50,2, 'Observações',ln=True)

#linha observação
# p = pdf.get_y()+50
p = pdf.get_y()+2
es = pdf.get_x()+1
pdf.set_line_width(1)
pdf.set_draw_color(0,0,0)
pdf.line(20,p, 125, p)

pdf.ln(3)

pdf.cell(12)
es = pdf.get_x()+1
p = pdf.get_y()+2
pdf.set_font('Arial', 'I', 12)

pdf.multi_cell(180, 10, 'Rachadura na contra verga -- sobre peso na laje -- novo pilar necessario -xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
footer(pdf)


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
    verificador = x%2
    
    if verificador == 0:
        pdf.image(img,10,p,w=180,h=90)
        pdf.ln(5)
        x+=1
        p+=145
        if  cont != test:
            print ('Break\n')
            pdf.add_page()
            p = pdf.get_y()+30
            cont+=1
            footer(pdf)
    else:    
        pdf.image(img,10,p,w=180,h=120)
        pdf.ln(5)
        x+=1
        p+=145
        cont+=1
        footer(pdf)
        # pdf.accept_page_break()
        
footer(pdf)


pdf.output('tuto2.pdf', 'F')