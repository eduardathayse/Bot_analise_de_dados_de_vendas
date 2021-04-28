"""
Robô para análise de dados de vendas.

Esse robô tem como passos:
• Abrir o navegador no drive onde esta o aquivo (planilha);
• Baixar a planilha;
• Calcular indicadores da planilha;
• Abrir o navegador na caixa de email;
• Digitar o email de destino o assunto e o corpo do email;
• Enviar email.
"""

# Bibliotecas essenciais para o funcionamento do robô
import pyautogui
import time
import pyperclip
import pandas as pd

class Robo_analise_dados:
    def __init__(self):
        """ Iniciando variáveis e primeiros comandos.
         """
        self.link_arquivo = 'https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym'
        self.link_email = 'https://mail.google.com/mail/u/0/#inbox?compose=new'
        self.navegador = "chrome"
        self.df = ''
        self.faturamento = ''
        self.qtd_produtos = ''
        self.email_destino = 'dudahmovies@gmail.com'
        self.assunto_email = 'Relatório de Vendas'
        self.corpo_email = """
        Presados, bom dia!

        O faturamento foi: R${}
        A quantidade de produtos foi de: {}
        
        Att,
        Eduarda Thayse
        """.format(self.faturamento,self.qtd_produtos)
        # Comandos iniciais
        time.sleep(2)
        pyautogui.alert('Iniciando...\n\nClick em OK para continuar e não mexa em nada.')
        time.sleep(2)
        # pyautogui.moveTo(0,0) # so para o mouse não atrapalhar
        

    def abrir_navegador(self):
        """ Abrindo o navegador chrome pelo executar. """
        time.sleep(3)
        pyautogui.hotkey('win','r')
        time.sleep(1)
        pyautogui.write(self.navegador)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(5)

    def drive_arquivo(self):
        """ Entrando no drive onde se encontra o arquivo (planilha). """
        time.sleep(3)
        pyperclip.copy(self.link_arquivo)
        pyautogui.hotkey('ctrl','v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(18)
        # pyautogui.PAUSE = 15

    def download_arquivo(self):
        """ Fazendo o download da planilha. """
        time.sleep(3)
        pyautogui.press('ctrl')
        time.sleep(2)
        pyautogui.press(['right', 'right', 'right'])
        time.sleep(1)
        pyautogui.hotkey('shift','f10') # atalho para o menu (Shift + F10,  Ctrl + Shift + F10)
        time.sleep(2)
        pyautogui.press('up')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(15)
        # chrome://settings/privacy     avançado   download   

    def planilha_dados(self):
        """ Atribuindo o planilha a um novo dataframe. """
        time.sleep(3)
        self.df = pd.read_excel(r'C:/Users/Eduarda/Downloads/Vendas - Dez.xlsx') # Caminho no no pc onde está salvo o arquivo Vendas.xlsx (pode mudar)
        print(self.df)

    def calculando_indiadores(self):
        """ Calculando os indicadores. 
        
        faturamento =  A soma do valor final 
        quantidade de produtos = A soma dos produtos
        """ 
        time.sleep(3)
        self.faturamento = self.df['Valor Final'].sum() # somando a coluna Valor Final
        self.qtd_produtos = self.df['Quantidade'].sum() # somando a coluna Quantidade
    
    def relatorio_via_email(self):
        """ Abrindo o gmail em uma nova guia no navegador.

        Realiaz os seguintes passos:
        # 1) ABRIR o gmail (novo e-mail);
        # 2) Digita o DESTINATÁRIO;
        # 3) Digita o ASSUNTO;
        # 4) Digita o CORPO DO E-MAIL usando os indicadores calculados;
        # 5) ENVIA E-mail;
        """
        time.sleep(3)
        pyautogui.hotkey('ctrl', 't')
        time.sleep(1)
        pyperclip.copy(self.link_email)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(35)
        pyautogui.write(self.email_destino)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.write(self.assunto_email)
        time.sleep(1)
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.write(self.corpo_email)
        time.sleep(1)
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('enter')
    
    def aviso_encerremanto(self):
        time.sleep(2)
        pyautogui.alert('Robô Finalizado!')

""" ======================= Script Principal ======================= """

try:
    robo = Robo_analise_dados() # instanciando o robô
    # chamando métodos do robô
    robo.abrir_navegador()
    robo.drive_arquivo()
    robo.download_arquivo()
    robo.planilha_dados()
    robo.calculando_indiadores()
    robo.relatorio_via_email()
    robo.aviso_encerremanto()
except Exception as erro:
    print('O erro foi ->> ',erro.__cause__)
    pyautogui.alert('Erro na execução!')