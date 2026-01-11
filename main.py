import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox
from tkinter.font import Font
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image, ImageTk
import qrcode
import pandas as pd
import xlrd
import datetime
import os
from io import BytesIO
import webbrowser

PRIMARY_COLOR = "#3498db"
SECONDARY_COLOR = "#2980b9"
BACKGROUND_COLOR = "#f5f5f5"
TEXT_COLOR = "#333333"
BUTTON_COLOR = "#2ecc71"
BUTTON_HOVER = "#27ae60"

MODELOS = {
    'ST': 'STREET',
    'CB': 'BOXY',
    'CL': 'CLÁSSICA',
    'ES': 'BABYLOOK',
    'IN': 'INFANTIL',
    'MG': 'M CARECA',
    'MC': 'M CANGURU',
    'RO': 'R OVERSIZED',
    'RB': 'R BOXY'
}

CORES = {
    "BR": "BRANCA",
    "PR": "PRETA",
    "MP": "MARMORIZADA",
    "CM": "MESCLA",
    "EC": "EST CHUMBO",
    "EM": "EST MARROM",
    "LL": "LILÁS",
    "MR": "MARROM",
    "OW": "OFF-WHITE",
    "VM": "VERDE MUSGO",
    "VR": "VERDE",
    "VE": "VERMELHA",
    "AM": "AZUL MARINHO",
    "AC": "AZUL CLARO"
}


def open_link(event):
    webbrowser.open_new("https://github.com/JoaoAHaupt")

class IzzyTagApp:
    def __init__(self, root):
        self.root = root
        self.setup_ui()

    def setup_ui(self):

        self.root.title('IzzyTag - Gerador de Etiquetas')
        self.root.resizable(False, False)
        self.root.geometry('700x450')
        self.root.configure(bg=BACKGROUND_COLOR)

        self.center_window()

        try:
            self.logo_image = Image.open("LOGO IZZY TAG.png")
            # Redimensionar mantendo a proporção (ajuste o tamanho conforme necessário)
            self.logo_image = self.logo_image.resize((200, 100), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(self.logo_image)

            logo_label = tk.Label(self.root, image=self.logo_photo, bg=BACKGROUND_COLOR)
            logo_label.pack(pady=(20, 10))
        except Exception as e:
            messagebox.showwarning("Aviso", f"Não foi possível carregar a logo: {str(e)}")

        self.title_font = Font(family="Helvetica", size=14, weight="bold")
        self.button_font = Font(family="Helvetica", size=10)


        main_frame = tk.Frame(self.root, bg=BACKGROUND_COLOR, padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)


        title_label = tk.Label(
            main_frame,
            text="IzzyTag - Gerador de Etiquetas IzzyPrint",
            font=self.title_font,
            bg=BACKGROUND_COLOR,
            fg=PRIMARY_COLOR
        )
        title_label.pack(pady=(0, 20))


        desc_label = tk.Label(
            main_frame,
            text=(
                "1. Selecione o arquivo Excel das etiquetas [.xls, .xlsx]\n"
                "2. Selecione a pasta onde será salvo."
            ),
            font=self.button_font,
            bg=BACKGROUND_COLOR,
            fg=TEXT_COLOR,
            wraplength=350
        )
        desc_label.pack(pady=(0, 30))



        self.open_button = tk.Button(
            main_frame,
            text='Selecionar Arquivo Excel',
            command=self.select_file,
            bg=BUTTON_COLOR,
            fg="white",
            activebackground=BUTTON_HOVER,
            activeforeground="white",
            font=self.button_font,
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        self.open_button.pack()

        footer_joao = tk.Label(
            main_frame,
            text="Por JoaoAHaupt",
            font=("Helvetica", 8, "underline"),
            bg=BACKGROUND_COLOR,
            fg="#0000EE",
            cursor="hand2"
        )

        footer_joao.pack(side=tk.BOTTOM, pady=(20, 0))
        footer_joao.bind("<Button-1>", open_link)  #

        footer_label = tk.Label(
            main_frame,
            text="© 2025 IzzyPrint - Todos os direitos reservados",
            font=("Helvetica", 8),
            bg=BACKGROUND_COLOR,
            fg="#777777"
        )
        footer_label.pack(side=tk.BOTTOM, pady=(20, 0))


        self.open_button.bind("<Enter>", lambda e: self.open_button.config(bg=BUTTON_HOVER))
        self.open_button.bind("<Leave>", lambda e: self.open_button.config(bg=BUTTON_COLOR))

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def sku_logic(self, sku):
        if not isinstance(sku, str) or sku == "":
            return "", "", "", ""

        parts = sku.split("-")
        tipo = produto = cor = tamanho = ""

        if len(parts) >= 4:
            _, _, _, info = parts

            if "FR" in info:
                tipo = "FR"
            elif "FC" in info:
                tipo = "FC"
            elif "CF" in info:
                tipo = "CF"

            info = info[4:]
            produto = MODELOS.get(info[0:2], "")
            cor = CORES.get(info[2:4], "")
            tamanho = info[4:]

        return tipo, produto, cor, tamanho

    def generate_pdf(self, df, path):
        c = canvas.Canvas(path)
        height = 302.36
        width = 377.95
        c.setPageSize((width, height))
        records = df.to_dict(orient="records")

        for row in records:
            font_name = "Helvetica"
            font_size = 10
            c.setFont(font_name, font_size)

            qr_size = 80
            text_height = height - 15
            qr_top = height - 100
            center_line = width / 2

            pedido_text = "Bling: " + str(row.get("Número do pedido", ""))
            pedido_width = c.stringWidth(pedido_text, font_name, font_size)

            x_text_left = (center_line/2) - (pedido_width/2)
            c.drawString(x_text_left, text_height, pedido_text)

            buffer1 = BytesIO()
            qr1 = qrcode.QRCode(box_size=2, border=2)
            qr1.add_data(pedido_text)
            qr1.make(fit=True)
            img_qr1 = qr1.make_image(fill_color="black", back_color="white")
            img_qr1.save(buffer1, format="PNG")
            buffer1.seek(0)
            qr_image1 = ImageReader(buffer1)

            x_qr1 = (center_line/2) - (qr_size/2)
            c.drawImage(qr_image1, x_qr1, qr_top, width=qr_size, height=qr_size)

            c.line(center_line, height, center_line, height-105)

            if row.get("Número gestão"):
                num_gest = str(row.get("Número gestão"))
                gestao_text = "Gestão: " + num_gest
                gestao_width = c.stringWidth(gestao_text, font_name, font_size)

                x_text_right = center_line + (center_line/2) - (gestao_width/2)
                c.drawString(x_text_right, text_height, gestao_text)

                buffer3 = BytesIO()
                qr3 = qrcode.QRCode(box_size=2, border=2)
                qr3.add_data(num_gest)
                qr3.make(fit=True)
                img_qr3 = qr3.make_image(fill_color="black", back_color="white")
                img_qr3.save(buffer3, format="PNG")
                buffer3.seek(0)
                qr_image3 = ImageReader(buffer3)

                x_qr3 = center_line + (center_line/2) - (qr_size/2)
                c.drawImage(qr_image3, x_qr3, qr_top, width=qr_size, height=qr_size)
                buffer3.close()

            buffer1.close()

            c.line(0, height - 105, width, height - 105)
            c.setFont("Helvetica-Bold", 14)
            c.drawString(10, height - 120, "Loja:" +" "+row.get("Nome da Loja", ""))
            c.drawString(10, height - 140, "Cliente:" +" "+row.get("Nome", ""))

            sku_value = row.get("Código (SKU)", "")
            if pd.notna(sku_value) and sku_value != "":
                sku_text = str(sku_value)
                buffer2 = BytesIO()
                qr2 = qrcode.QRCode(box_size=2, border=2)
                qr2.add_data(sku_text)
                qr2.make(fit=True)
                img_qr2 = qr2.make_image(fill_color="black", back_color="white")
                img_qr2.save(buffer2, format="PNG")
                buffer2.seek(0)
                qr_image2 = ImageReader(buffer2)

                qr2_size = 120
                c.drawImage(qr_image2, (width - qr2_size), height - 300, width=qr2_size, height=qr2_size)
                c.drawString(10, height - 290, sku_text)

                tipo, produto, cor, tamanho = self.sku_logic(sku_text)
                c.setFont("Helvetica-Bold", 18)
                c.drawString(10, height - 250, tamanho)
                c.drawString(10, height - 230, cor)
                c.drawString(10, height - 210, produto)

                c.setFont("Helvetica-Bold", 28)
                c.drawString(width/2-24, height - 230, tipo)
                buffer2.close()
            else:
                c.setFont("Helvetica-Bold", 23)
                sem_sku_text = "SEM SKU"
                text_width = c.stringWidth(sem_sku_text, "Helvetica-Bold", 23)
                c.drawString((width - text_width) / 2, height - 220, sem_sku_text)
                c.setFont(font_name, 15)

            custom_style = ParagraphStyle(
                'Custom',
                parent=getSampleStyleSheet()['Normal'],
                fontName='Helvetica',
                fontSize=12,
                leading=13,
            )

            descricao = f"{row.get('Descrição', '')}<br/>"
            p = Paragraph(descricao, custom_style)
            p.wrapOn(c, width - 20, 80)
            p.drawOn(c, 10, height - 180)

            c.showPage()

        c.save()

    def select_file(self):
        filetypes = (
            ("Excel files", "*.xls *.xlsx"),
            ("All files", "*.*")
        )

        filename = fd.askopenfilename(
            title='Selecione um arquivo Excel',
            initialdir='/',
            filetypes=filetypes
        )

        if not filename:
            return

        save_path = fd.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Salvar PDF como",
            initialfile=f"Etiquetas_{datetime.datetime.now().strftime('%d-%m-%Y_%H:%M')}.pdf"
        )

        if not save_path:
            messagebox.showinfo(
                title='Cancelado',
                message='Operação cancelada. Nenhum PDF será gerado.'
            )
            return

        try:
            self.open_button.config(state=tk.DISABLED, text="Processando...")
            self.root.update()

            ext = os.path.splitext(filename)[1].lower()
            if ext == ".xls":
                workbook = xlrd.open_workbook(filename, ignore_workbook_corruption=True)
                df = pd.read_excel(workbook)
            elif ext == ".xlsx":
                df = pd.read_excel(filename)
            else:
                raise ValueError("Formato de arquivo não suportado")

            self.generate_pdf(df, save_path)

            messagebox.showinfo(
                title='Sucesso',
                message=f'PDF gerado com sucesso em:\n{save_path}'
            )

        except Exception as e:
            messagebox.showerror(
                title='Erro',
                message=f'Ocorreu um erro:\n{str(e)}'
            )
        finally:
            self.open_button.config(state=tk.NORMAL, text="Selecionar Arquivo Excel")

def main():
    root = tk.Tk()
    app = IzzyTagApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
