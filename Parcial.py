#    Copyright (C) 2022  Igor Pereira Formighieri <igorpereira1069@gmail.com>


#    Este programa é um software livre: você pode redistribuí-lo e/ou
#    modificá-lo sob os termos da Licença Pública Geral GNU, conforme
#    publicado pela Free Software Foundation, seja a versão 3 da Licença
#    ou (a seu critério) qualquer versão posterior.

#    Este programa é distribuído na esperança de que seja útil,
#    mas SEM QUALQUER GARANTIA; sem a garantia implícita de
#    COMERCIALIZAÇÃO OU ADEQUAÇÃO A UM DETERMINADO PROPÓSITO. Veja a
#    Licença Pública Geral GNU para obter mais detalhes.

#    Você deve ter recebido uma cópia da Licença Pública Geral GNU
#    junto com este programa. Se não, veja <https://www.gnu.org/licenses/>.


from asyncio.windows_events import NULL
from fileinput import close
from webbrowser import get
from openpyxl import Workbook, load_workbook
from datetime import datetime
from tkinter import *

from tkinter import filedialog as dlg
from tkinter import Tk
from tkinter.filedialog import askopenfilename


from PIL import ImageTk, Image
import os
import sys


    

class janela_class(Tk):
    def __init__(self):
        super().__init__()
        super().overrideredirect(True)
        
    
        self._offsetx = 0
        self._offsety = 0
        super().bind("<Button-1>" ,self.clickwin)
        super().bind("<B1-Motion>", self.dragwin)

    def dragwin(self,event):
        x = super().winfo_pointerx() - self._offsetx
        y = super().winfo_pointery() - self._offsety
        super().geometry(f"+{x}+{y}")

    def clickwin(self,event):
        self._offsetx = super().winfo_pointerx() - super().winfo_rootx()
        self._offsety = super().winfo_pointery() - super().winfo_rooty()

janela = janela_class()




def calcular_tabela():
    
    Tk().withdraw() # Isto torna oculto a janela principal
    abrir_janela = False
    try:
        
        planilha = load_workbook(askopenfilename())
        aba_ativa = planilha.active
        
        botao_calcular['state'] = DISABLED
        botao_calcular['bg'] = "white"
        
        ALTO_GIRO = 0
        BAZAR = 0
        DIVERSOS = 0
        DPH = 0
        FLV = 0
        LATICINIOS = 0
        LIQUIDA = 0
        PERECIVEL1 = 0
        PERECIVEL2 = 0
        PERECIVEL2B = 0
        PERECIVEL3 = 0
        SECA_DOCE = 0
        SECA_SALGADA = 0
        SECA_SALGADA2 = 0
        TOTAL = 0
        PARCIAL = 0
        HORA_DATA = datetime.today().strftime('%H:%M – %d/%m/%Y')

        for celula in aba_ativa["C"]:
            TOTAL = TOTAL + 1
            if celula.value == "ALTO GIRO":
                ALTO_GIRO = ALTO_GIRO + 1
            if celula.value == "BAZAR":
                BAZAR = BAZAR + 1
            if celula.value == "DPH":
                DPH = DPH + 1
            if celula.value == "DIVERSOS":
                DIVERSOS = DIVERSOS + 1
            if celula.value == "FLV":
                FLV = FLV + 1
            if celula.value == "LATICINIOS 1":
                LATICINIOS = LATICINIOS + 1
            if celula.value == "LIQUIDA":
                LIQUIDA = LIQUIDA + 1
            if celula.value == "PERECIVEL 1":
                PERECIVEL1 = PERECIVEL1 + 1
            if celula.value == "PERECIVEL 2":
                PERECIVEL2 = PERECIVEL2 + 1
            if celula.value == "PERECIVEL 2 B":
                PERECIVEL2B = PERECIVEL2B + 1
            if celula.value == "PERECIVEL 3":
                PERECIVEL3 = PERECIVEL3 + 1
            if celula.value == "SECA DOCE":
                SECA_DOCE = SECA_DOCE + 1
            if celula.value == "SECA SALGADA":
                SECA_SALGADA = SECA_SALGADA + 1
            if celula.value == "SECA SALGADA 2":
                SECA_SALGADA2 = SECA_SALGADA2 + 1
        
        def parcial():
            PARCIAL = round(TOTAL/int(SKU.get())*100,2)
            return str(PARCIAL) + "%"
        PARCIAL = parcial()

        texto_parcial_qtd1 = Label(janela, text=ALTO_GIRO, bg="#2a6099", fg="white")
        texto_parcial_qtd2 = Label(janela, text=BAZAR, bg="#2a6099", fg="white")
        texto_parcial_qtd3 = Label(janela, text=DIVERSOS, bg="#2a6099", fg="white")
        texto_parcial_qtd4 = Label(janela, text=DPH, bg="#2a6099", fg="white")
        texto_parcial_qtd5 = Label(janela, text=FLV, bg="#2a6099", fg="white")
        texto_parcial_qtd6 = Label(janela, text=LATICINIOS, bg="#2a6099", fg="white")
        texto_parcial_qtd7 = Label(janela, text=LIQUIDA, bg="#2a6099", fg="white")
        texto_parcial_qtd8 = Label(janela, text=PERECIVEL1, bg="#2a6099", fg="white")
        texto_parcial_qtd9 = Label(janela, text=PERECIVEL2, bg="#2a6099", fg="white")
        texto_parcial_qtd10 = Label(janela, text=PERECIVEL2B, bg="#2a6099", fg="white")
        texto_parcial_qtd11 = Label(janela, text=PERECIVEL3, bg="#2a6099", fg="white")
        texto_parcial_qtd12 = Label(janela, text=SECA_DOCE, bg="#2a6099", fg="white")
        texto_parcial_qtd13 = Label(janela, text=SECA_SALGADA, bg="#2a6099", fg="white")
        texto_parcial_qtd14 = Label(janela, text=SECA_SALGADA2, bg="#2a6099", fg="white")
        texto_parcial_qtd_total = Label(janela, text=TOTAL, bg="#5983b0", fg="white")
        texto_parcial = Label(janela, text=PARCIAL, bg="#2a6099", fg="white")
        #texto_sku = Label(janela, text=SKU, bg="#2a6099", fg="white")
        texto_hora_data = Label(janela, text=HORA_DATA, bg="#2a6099", fg="white")
        

        texto_parcial_qtd1.config(font=('Arial Black',8))
        texto_parcial_qtd2.config(font=('Arial Black',8))
        texto_parcial_qtd3.config(font=('Arial Black',8))
        texto_parcial_qtd4.config(font=('Arial Black',8))
        texto_parcial_qtd5.config(font=('Arial Black',8))
        texto_parcial_qtd6.config(font=('Arial Black',8))
        texto_parcial_qtd7.config(font=('Arial Black',8))
        texto_parcial_qtd8.config(font=('Arial Black',8))
        texto_parcial_qtd9.config(font=('Arial Black',8))
        texto_parcial_qtd10.config(font=('Arial Black',8))
        texto_parcial_qtd11.config(font=('Arial Black',8))
        texto_parcial_qtd12.config(font=('Arial Black',8))
        texto_parcial_qtd13.config(font=('Arial Black',8))
        texto_parcial_qtd14.config(font=('Arial Black',9))
        texto_parcial_qtd_total.config(font=('Arial Black',9))
        texto_parcial.config(font=('Arial Black',9))
        #texto_sku.config(font=('Arial Black',9))
        texto_hora_data.config(font=('Arial Black',20))

        texto_parcial_qtd1.place(x=390, y=180)
        texto_parcial_qtd2.place(x=390, y=203)
        texto_parcial_qtd3.place(x=390, y=227)
        texto_parcial_qtd4.place(x=390, y=250)
        texto_parcial_qtd5.place(x=390, y=273)
        texto_parcial_qtd6.place(x=390, y=298)
        texto_parcial_qtd7.place(x=390, y=324)
        texto_parcial_qtd8.place(x=390, y=350)
        texto_parcial_qtd9.place(x=390, y=378)
        texto_parcial_qtd10.place(x=390, y=403)
        texto_parcial_qtd11.place(x=390, y=430)
        texto_parcial_qtd12.place(x=390, y=455)
        texto_parcial_qtd13.place(x=390, y=478)
        texto_parcial_qtd14.place(x=390, y=501)
        texto_parcial_qtd_total.place(x=390, y=525)
        texto_parcial.place(x=602, y=324)
        #texto_sku.place(x=809, y=324)
        texto_hora_data.place(x=582, y=378)
        
    except Exception:
        janela.destroy()
        #abrir_janela = True
        sys.exit()

    finally:

        #if abrir_janela == True:
        #    os.system(sys.argv[0])
        #else:
        def limpar():
            texto_parcial_qtd1.destroy()
            texto_parcial_qtd2.destroy()
            texto_parcial_qtd3.destroy()
            texto_parcial_qtd4.destroy()
            texto_parcial_qtd5.destroy()
            texto_parcial_qtd6.destroy()
            texto_parcial_qtd7.destroy()
            texto_parcial_qtd8.destroy()
            texto_parcial_qtd9.destroy()
            texto_parcial_qtd10.destroy()
            texto_parcial_qtd11.destroy()
            texto_parcial_qtd12.destroy()
            texto_parcial_qtd13.destroy()
            texto_parcial_qtd14.destroy()
            texto_parcial.destroy()
            texto_hora_data.destroy()
            texto_parcial_qtd_total.destroy()
            texto_hora_data.destroy()
            #texto_sku.destroy()
            botao_calcular['state'] = NORMAL
            botao_calcular['bg'] = "#2a6099"
            botao_limpar['state'] = DISABLED
            botao_limpar['bg'] = "white"

        
        botao_limpar = Button(janela, text="Limpar", command=limpar, bg="#2a6099", fg="white")
        botao_limpar.place(x=100, y=3)




janela.title("Parcial Da Auditoria")
janela.geometry("925x552")
img = ImageTk.PhotoImage(Image.open("img/background.png"))
imglabel = Label(janela, image=img).place(x=0, y=30)

SKU = Entry(janela, width=5, bg="#2a6099", fg="white")
SKU.place(x=805, y=324)
SKU.config(font=('Arial Black',9))

def calcular():
    
    if SKU.get():
        calcular_tabela()
    else:
        pass
    
botao_calcular = Button(janela, text="Calcular Parcial", command=calcular, bg="#2a6099", fg="white")
botao_calcular.place(x=5, y=3)

botao_fechar = Button(janela, text="X", command=sys.exit, bg="#d92828", width= 5, fg="white")
botao_fechar.place(x=870, y=0)
botao_fechar.config(font=('Arial Black',10))



janela.mainloop()

sys.exit()
