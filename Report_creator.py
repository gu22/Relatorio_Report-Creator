# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 19:09:02 2021

@author: gusan
"""

from fpdf import FPDF

H = 297
W = 210




pdf= FPDF('P','mm','A4')
pdf.add_page()


#Linha lateral
pdf.set_line_width(5)
pdf.set_draw_color(0,45,0)
pdf.line(10, 10, 10, 290)

#Header
pdf.set_font('Arial', 'B', 16)
pdf.cell(12)
pdf.cell(50, 10, 'Relatório - Ordem de Serviço')
pdf.ln(2)
pdf.set_line_width(1)
pdf.set_draw_color(0,0,0)
pdf.line(20, 18, 110, 18)


pdf.output('tuto2.pdf', 'F')