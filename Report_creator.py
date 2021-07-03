# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 19:09:02 2021

@author: gusan
"""

from fpdf import FPDF
import easygui
import math
import pandas as pd
from unidecode import unidecode

from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox
import sys
import os

import configparser


config = configparser.ConfigParser()
config.read('Config.ini')
design = config['design']
default = config['DEFAULT']


rgb = (design['Cor_linhalateral']).split(',')
rgb = [int(x)for x in rgb]
r,g,b = rgb[0],rgb[1],rgb[2]

H_logo = int(design['Logo_altura'])
W_logo = int(design['Logo_comprimento'])

titulo = (default['Titulo'])
subtitulo = (default['Subtitulo'])
subtitulo_2n = (default['Subnivel_2'])


'''GUI Otimizada para  Usuario'''

class Ui(QtWidgets.QMainWindow):
    
    ''' Variaveis Globais '''
    #titulo, subtitulo, subtitulo_2n,r,g,b,base_open,files,H_logo,W_logo
    
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('Main.ui', self)
        
        self.setFixedSize(316, 357)
        
        self.Status_dados.setText('')
        self.Status_imagens.setText('')
        self.Status_emissao.setText('')
        
        self.Bardados.setValue(0)
        self.Barimagens.setValue(0)
        
        
        self.Bdados.clicked.connect(self.open_dados)
        self.Bimagens.clicked.connect(self.open_imagens)
        self.Bemissao.clicked.connect(self.r_emissao)
        
        self.menuConfig.triggered.connect(self.pp)
        
        self.show()
        
        



    def pp(self):
        # print ('OK')
        msgBox = QMessageBox()
        msgBox.setText("Para editar as configurações, por favor, utilize o bloco de notas/Notepad")
        msgBox.exec()
        os.system('tt.txt')


    def open_dados(self):
        global base_open
        base_open = easygui.fileopenbox()
        self.Status_dados.setText('Dados OK')
        self.Bardados.setValue(100)

    def open_imagens(self):
        global files
        
        files = easygui.fileopenbox(multiple=True)
        self.Status_imagens.setText('Imagens OK ')
        self.Barimagens.setValue(100)
        


    def r_emissao(self):
        global files , base_open,r,g,b,H_logo,W_logo,titulo, subtitulo, subtitulo_2n
        
        ''' GUI para leitura de arquivo '''
        
        # base_open = easygui.fileopenbox()
        
        
        
        '''Leitura dos Arquivos'''
        
        itens = ('ID','N° Ordem','Unidade','Pavimento','Requisição','')
        dados = ('00','1234','Santana','T','Trocar Janela quebrada','')
        
        # base = pd.read_excel(base_open,engine='openpyxl')
        # head_file = list(base)
        
        # index  = len(base.index)
        # itens = head_file
        # dados = []
        
        # # for i in range(index)
        
        # for i in range(len(itens)):
        #     if  head_file[i] == ('Observações'):
        #         obs = (str(base.iloc[0][i]))
        #         itens.remove('Observações')
        #     elif head_file[i] == ('N°  Ordem'):
        #         name_file = (str(base.iloc[0][i]))
        #         dados.append(str(base.iloc[0][i]))
        #     else:
        #         dados.append(str(base.iloc[0][i]))
        
        
        
        
        
        obs = 'Teste'
        
        
        
        ''' Configuração da pagina '''
        
        H = 297
        W = 210
        
        pdf= FPDF('P','mm','A4')
        pdf.set_right_margin(10)
        # pdf.page_no()
        pdf.alias_nb_pages()
        pdf.set_auto_page_break(True,10)
        
        pdf.add_page()
        
        
        ''' Linha lateral'''
        pdf.set_line_width(5)
        pdf.set_draw_color(r,g,b)
        pdf.line(10, 50, 10, 290)
        
        ''' Header '''
        pdf.image(files[0],5,10,w=W_logo,h=H_logo)
        
        pdf.set_title('Report')
        pdf.set_font('Arial', 'B', 20)
        # pdf.cell(50)
        pdf.cell(0, 8, titulo,ln=True,align="C")
        # pdf.cell(55)
        pdf.set_font('Arial', 'B', 18)
        pdf.cell(0, 8, subtitulo,ln=True,align="C")
        # pdf.cell(60)
        pdf.ln(7)
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 0, subtitulo_2n,align="C")
        
        pdf.ln(5)
        pdf.set_font('courier', 'I', 14)
        pdf.cell(150)
        pdf.cell(10, 12, 'N° ')
        pdf.cell(50, 12, 'xxxx-xxx-xx ')
        pdf.ln(-5)
        # pdf.set_line_width(1)
        # pdf.set_draw_color(0,0,0)
        # pdf.line(20, 19, 125, 19)
        
        c = 0
        
        ''' Footer'''
        
        def footer(self):
            # Position at 1.5 cm from bottom
                self.set_y(-15)
                # helvetica italic 8
                self.set_font('helvetica', 'I', 8)
                # Page number
                self.cell(0, 0, 'Página ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')
                self.ln(5)
                self.cell(0, 0, 'BBXYBXYUBXYBXYUBXU',0,0,'C')
        
        # pdf.cell(0,)
        
        '''Informações'''
        
        
        #INFORMAÇÕES INICIAIS
        pdf.set_font('Arial', 'B', 16)
        p = pdf.get_y()+10
        pdf.cell(12)
        pdf.cell(50, p, f'Unidade Requisitante: ',1)
        
        es = pdf.get_x()+3
        pdf.set_font('Arial', 'I', 12)
        pdf.cell(55, p, 'TETSTSTS')
        
        es = pdf.get_x()+10
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(40, p, 'Data: ')
        
        es = pdf.get_x()-20
        pdf.set_font('Arial', 'I', 12)
        pdf.cell(0, p, '20212021')
        
        
        pdf.ln(12)
        
        
        
        
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
        
        pdf.multi_cell(180, 10, f'{obs}')
        footer(pdf)
        
        
        '''Imagens'''
        
        
        pdf.add_page()
        pdf.image(files[0],5,10,w=W_logo,h=H_logo)
        
        
        # pdf.ln(10)
        # pdf.cell(12)
        # p = pdf.get_y()+40
        pdf.set_font('Arial', 'B', 16)
        # pdf.cell(20)
        pdf.cell(0, 10, 'Imagens',align='C')
        
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
                    print ('PageBreak\n')
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
        
        
        pdf.output(f'teste_r.pdf', 'F')
        
        self.Status_dados.setText('')
        self.Status_imagens.setText('')
        self.Status_emissao.setText('Emissão Concluida')
        
        self.Bardados.setValue(0)
        self.Barimagens.setValue(0)
        

app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()