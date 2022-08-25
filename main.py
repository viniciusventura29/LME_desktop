from PyQt5 import uic, QtWidgets,QtCore, QtGui
import matplotlib.pyplot as plt
import pandas as pd
from datetime import date
import win32com.client as cli

app = QtWidgets.QApplication([])

tela_home = uic.loadUi('janelas/tela_home.ui')
tela_opcoes = uic.loadUi('janelas/tela_opcoes.ui')
tela_variacao = uic.loadUi('janelas/tela_variacao_metal.ui')
tela_clientes = uic.loadUi('janelas/tela_escolha_cliente.ui')
tela_tamarana = uic.loadUi('janelas/tela_tamarana.ui')
tela_atualizacao = uic.loadUi('janelas/tela_atualizacao.ui')
tela_aviso_atualizacao = uic.loadUi('janelas/tela_atencao_atualizacao.ui')
tela_aviso_med_mensal = uic.loadUi('janelas/tela_aviso_med_mensalidade.ui')

fonte = pd.read_excel('planilha/LME_fonte_de_dados.xlsx')
media_mensal = pd.read_excel('planilha/LME_media_mensal.xlsx')
table_dolar = pd.read_excel('planilha/LME_fonte_de_dados.xlsx', sheet_name='dolar')


class Main:
    def exibicao_home():
        tela_opcoes.close()
        tela_variacao.close()
        tela_home.show()

    def exibicao_opcoes():
        tela_home.close()
        tela_clientes.close()
        tela_variacao.close()
        tela_opcoes.show()

    def exibicao_variacao():
        tela_opcoes.close()
        tela_variacao.show()

    def exibicao_clientes():
        tela_opcoes.close()
        tela_clientes.show()

    def exibicao_tamarana():
        tela_tamarana.show()

    def chamada_grafico(num, metal):
        grafico.exibicao(num, metal)

    def chamada_cliente():
        cliente.definicao_de_parametros()
        cliente.preco_base()
        cliente.media_3_meses()
        cliente.variacao_porcentual()
        cliente.gatilho()

    def atualizar_cliente():
        cliente.atualizar()

    def close_tela_atualizacao():
        tela_atualizacao.close()
        tela_tamarana.preco_base.clear()
        tela_tamarana.media_3_meses.clear()
        tela_tamarana.variacao_porcentual.clear()
        tela_tamarana.gatilho.clear()

    def aviso_atualizacao_atualizar():
        tela_tamarana.close()
        tela_aviso_atualizacao.close()
        tela_tamarana.preco_base.clear()
        tela_tamarana.media_3_meses.clear()
        tela_tamarana.variacao_porcentual.clear()
        tela_tamarana.gatilho.clear()

    def aviso_atualizacao_cancelar():
        tela_aviso_atualizacao.close()

    def atualizar_excel():
        atualizar.mes_atual()
        atualizar.adicao_valor()
        atualizar.conversion()
        atualizar.adicionar()
        tela_aviso_med_mensal.show()

    def close_aviso_med():
        tela_aviso_med_mensal.close()
        
    def getExcel():
        outlook = cli.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")

        acc = namespace.Folders['Campinas.ETS@br.bosch.com']
        inbox = acc.Folders('Inbox')

        all_inbox = inbox.Items

        for msg in all_inbox:
            if msg.Class==43:
                if msg.SenderEmailType=='EX':
                    pass
                else:
                    if msg.SenderEmailAddress == "viniciusventura29@icloud.com":
                        #chamar tela de notificação
                        print(msg)
                        msg.Move(acc.Folders('Teste'))
                            
        

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(248, 268)
        Form.setWindowIcon(QtGui.QIcon("img/calendar-service.svg"))
        Form.setMaximumHeight(268)
        Form.setMaximumWidth(248)
        Form.setStyleSheet("background-color:#fff;")
        self.icon_info = QtWidgets.QLabel(Form)
        self.icon_info.setGeometry(QtCore.QRect(90, 5, 71, 71))
        self.icon_info.setText("")
        self.icon_info.setPixmap(QtGui.QPixmap("img/alert-warning-filled.svg"))
        self.icon_info.setScaledContents(True)
        self.icon_info.setObjectName("icon_info")
        self.label_atencao = QtWidgets.QLabel(Form)
        self.label_atencao.setGeometry(QtCore.QRect(-90, 85, 431, 41))
        font = QtGui.QFont()
        font.setFamily("Bosch Sans Bold")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_atencao.setFont(font)
        self.label_atencao.setStyleSheet("color:#007BC0;letter-spacing:2px")
        self.label_atencao.setScaledContents(False)
        self.label_atencao.setAlignment(QtCore.Qt.AlignCenter)
        self.label_atencao.setObjectName("label_atencao")
        self.button_ok = QtWidgets.QPushButton(Form)
        self.button_ok.setGeometry(QtCore.QRect(70, 195, 111, 31))
        self.button_ok.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.button_ok.setStyleSheet("background-color:#007BC0;color:#fff;border-radius:7px")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("img/checkmark-bold (1).svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button_ok.setIcon(icon)
        self.button_ok.setIconSize(QtCore.QSize(18, 18))
        self.button_ok.setAutoExclusive(False)
        self.button_ok.setDefault(False)
        self.button_ok.setObjectName("button_ok")

        self.button_ok.clicked.connect(Form.close)

        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(0, 260, 251, 151))
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("img/Bosch-Supergraphic.svg"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.label_atencao_2 = QtWidgets.QLabel(Form)
        self.label_atencao_2.setGeometry(QtCore.QRect(-90, 135, 431, 31))
        font = QtGui.QFont()
        font.setFamily("Bosch Sans Medium")
        font.setPointSize(8)
        font.setBold(False)
        font.setWeight(50)
        self.label_atencao_2.setFont(font)
        self.label_atencao_2.setStyleSheet("color:#007BC0;letter-spacing:2px")
        self.label_atencao_2.setScaledContents(False)
        self.label_atencao_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_atencao_2.setObjectName("label_atencao_2")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle("Lembrete")
        self.label_atencao.setText(_translate("Form", "<html><head/><body><p align=\"center\">CONTRATO <br/>DESATUALIZADO</p></body></html>"))
        self.button_ok.setText(_translate("Form", "OK"))
        self.label_atencao_2.setText(_translate("Form", "<html><head/><body><p align=\"center\">Abra o executavel<br/>para atualiza-lo</p></body></html>"))

class Grafico:
    def __init__(self):
        pass

    def exibicao(self, num, escolha):
        print(escolha)
        X = list(media_mensal.iloc[:, 0])
        Y = list(media_mensal.iloc[:, num])

        plt.bar(X, Y)
        plt.title(escolha)
        plt.xlabel('Mês')
        plt.ylabel('Preço')
        plt.show()

class Atualizar:
    def __init__(self):
        self.med_cobre = fonte['Cobre U$/t'].iloc[-1]
        self.med_zinco = fonte['Zinco U$/t'].iloc[-1]
        self.med_aluminio = fonte['Alumínio U$/t'].iloc[-1]
        self.med_chumbo = fonte['Chumbo U$/t'].iloc[-1]
        self.med_estanho = fonte['Estanho U$/t'].iloc[-1]
        self.med_niquel = fonte['Níquel U$/t'].iloc[-1]
        self.values = [self.med_cobre, self.med_zinco, self.med_aluminio, self.med_chumbo, self.med_estanho, self.med_niquel]
        self.month = date.today().strftime('%m')
        self.numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
        self.meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        self.price_dolar = table_dolar.iloc[0, 1]

    def adicionar(self):
        self.values.insert(0, self.mes_escrit)
        print(self.values)
        media_mensal.loc[len(media_mensal) + 1] = self.values
        media_mensal.to_excel('planilha/LME_media_mensal.xlsx', index = False)

    def mes_atual(self):
        cont = 0
        self.month = int(self.month)
        for i in self.numbers:
            if self.month == i:
                self.mes_escrit = self.meses[cont]
            cont = cont + 1

    def adicao_valor(self):
        aux = self.values.copy()
        self.values.clear()
        for i in aux:
            valor = list(i)
            for j in range(3):
                valor.pop(-1)
            self.values.append(''.join(valor))

    def conversion(self):
        aux = self.values.copy()
        self.values.clear()
        for i in aux:
            i = i.replace(',','.')
            i = float(i)
            i = i * self.price_dolar
            i = round(i, 3)
            self.values.append(i)

class Cliente:
    def __init__(self):
        pass

    def definicao_de_parametros(self):
        self.material_selecionado = tela_tamarana.comboBox_material.currentText()
        self.ultimo_mes = media_mensal[self.material_selecionado].iloc[-1]
        self.penultimo_mes = media_mensal[self.material_selecionado].iloc[-2]
        self.antipenultimo_mes = media_mensal[self.material_selecionado].iloc[-3]

    def preco_base(self):
        self.base_teste = 15.698
        tela_tamarana.preco_base.setText(str(self.base_teste))

    def media_3_meses(self):
        self.media = (self.antipenultimo_mes + self.penultimo_mes + self.ultimo_mes) / 3
        media_string = '%.2f' % (self.media)
        tela_tamarana.media_3_meses.setText(media_string)

    def variacao_porcentual(self):
        self.porcent_med = (self.media * 100) / self.base_teste
        self.var_porcent = 100 - self.porcent_med
        var_porcent_string = '%.0f' % (self.var_porcent) + '%'
        tela_tamarana.variacao_porcentual.setText(var_porcent_string)

    def gatilho(self):
        if self.var_porcent >= 50:
            tela_tamarana.gatilho.setText('Atingiu o gatilho!')
            self.var_gatilho = True
        else:
            tela_tamarana.gatilho.setText('Não atingiu o gatilho!')
            self.var_gatilho = False

    def atualizar(self):
        if self.var_gatilho == True:
            tela_tamarana.close()
            tela_atualizacao.show()
        else:
            tela_aviso_atualizacao.show()

grafico = Grafico()
cliente = Cliente()
atualizar = Atualizar()
ui_form = Ui_Form()


#BTN TELA HOME
tela_home.Button_avancar.clicked.connect(Main.exibicao_opcoes)

#BTN TELA OPCOES
tela_opcoes.Button_voltar.clicked.connect(Main.exibicao_home)
tela_opcoes.Button_variacao.clicked.connect(Main.exibicao_variacao)
tela_opcoes.Button_atualizar_preco.clicked.connect(Main.exibicao_clientes)
tela_opcoes.Button_atualizar_exc.clicked.connect(Main.atualizar_excel)

#BTN TELA VARIACOES
tela_variacao.Button_home.clicked.connect(Main.exibicao_opcoes)
tela_variacao.btn_cobre.clicked.connect(lambda: Main.chamada_grafico(1, 'Cobre'))
tela_variacao.btn_zinco.clicked.connect(lambda: Main.chamada_grafico(2, 'Zinco'))
tela_variacao.btn_aluminio.clicked.connect(lambda: Main.chamada_grafico(3, 'Aluminio'))
tela_variacao.btn_chumbo.clicked.connect(lambda: Main.chamada_grafico(4, 'Chumbo'))
tela_variacao.btn_estanho.clicked.connect(lambda: Main.chamada_grafico(5, 'Estanho'))
tela_variacao.btn_niquel.clicked.connect(lambda: Main.chamada_grafico(6, 'Niquel'))

#BTN TELA CLIENTES
tela_clientes.Button_home.clicked.connect(Main.exibicao_opcoes)
tela_clientes.Button_tamarana.clicked.connect(Main.exibicao_tamarana)

#BTN TELA TAMARANA
tela_tamarana.comboBox_material.addItem('Chumbo')
tela_tamarana.comboBox_material.addItem('Zinco')
tela_tamarana.comboBox_material.addItem('Aluminio')
tela_tamarana.comboBox_material.addItem('Estanho')
tela_tamarana.Button_pesquisa.clicked.connect(Main.chamada_cliente)
tela_tamarana.Button_atualizar.clicked.connect(Main.atualizar_cliente)

#BTN TELA ATUALIZACAO
tela_atualizacao.Button_ok.clicked.connect(Main.close_tela_atualizacao)

#BTN TELA AVISO ATUALIZACAO
tela_aviso_atualizacao.Button_atualizar.clicked.connect(Main.aviso_atualizacao_atualizar)
tela_aviso_atualizacao.Button_cancelar.clicked.connect(Main.aviso_atualizacao_cancelar)

#BTN TELA AVISO ATUALIZACAO DE MEDIA MENSAL
tela_aviso_med_mensal.Button_ok.clicked.connect(Main.close_aviso_med)


tela_home.show()
app.exec()

