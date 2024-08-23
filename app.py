from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QListWidget, QTabWidget, QCompleter
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import yagmail
import os
import sys
import json
import shutil
import mysql.connector


class EmailApp(QWidget):
    def __init__(self):
        super().__init__()

        self.caminhos_anexo = []  # Lista para armazenar os caminhos dos arquivos
        self.config_path = "config.json"  # Caminho do arquivo de configuração
        self.gmail_usuario = ""  # Armazenar o Gmail do usuário
        self.gmail_senha = ""  # Armazenar a senha do Gmail do usuário
        self.db_host = ""  # Host do banco de dados padrão
        self.db_usuario = ""  # Usuário do banco de dados padrão
        self.db_senha = ""  # Senha do banco de dados padrão
        self.db_nome = ""  # Nome do banco de dados padrão
        self.clientes = []  # Lista de clientes carregados
        self.pasta_origem = ""  # Pasta padrão de origem dos PDFs
        self.pasta_destino = ""  # Pasta padrão de destino dos PDFs

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Envio de E-mail com Anexo")
        self.setGeometry(100, 100, 800, 600)

        # Definir ícone da janela com a imagem da logo
        self.setWindowIcon(QIcon("logo.png"))  # Substitua "logo.png" pelo nome do seu arquivo de logo

        # Configuração do TabWidget
        self.tabs = QTabWidget()  # Alteração aqui
        tab_email = QWidget()
        tab_config = QWidget()
        tab_orcamentos = QWidget()

        self.tabs.addTab(tab_email, "Enviar E-mail")
        self.tabs.addTab(tab_config, "Configurações")
        self.tabs.addTab(tab_orcamentos, "Orçamentos")

        # Layout principal
        layout_principal = QVBoxLayout(self)
        layout_principal.addWidget(self.tabs)

        # Layout da aba de envio de e-mails
        self.init_email_tab(tab_email)

        # Layout da aba de configurações
        self.init_config_tab(tab_config)

        # Layout da aba de orçamentos
        self.init_orcamentos_tab(tab_orcamentos)

        self.setLayout(layout_principal)

        # Carregar configurações e clientes após a inicialização dos widgets
        self.load_config()
        self.load_clientes()
        self.configurar_autocompletar()
        self.atualizar_orcamentos()

    def init_email_tab(self, tab_email):
        layout_principal = QVBoxLayout()

        # Campo de pesquisa de cliente
        layout_cliente = QHBoxLayout()
        lbl_cliente = QLabel("Cliente:")
        self.entry_cliente = QLineEdit()
        self.entry_cliente.textChanged.connect(self.preencher_email_cliente)
        layout_cliente.addWidget(lbl_cliente)
        layout_cliente.addWidget(self.entry_cliente)

        # Botão para atualizar clientes
        btn_atualizar_clientes = QPushButton("Atualizar Clientes", clicked=self.load_clientes)
        layout_cliente.addWidget(btn_atualizar_clientes)

        layout_principal.addLayout(layout_cliente)

        # Campo de e-mail
        layout_email = QHBoxLayout()
        lbl_email = QLabel("E-mail:")
        self.entry_email = QLineEdit()
        layout_email.addWidget(lbl_email)
        layout_email.addWidget(self.entry_email)
        layout_principal.addLayout(layout_email)

        # Campo de Assunto
        layout_assunto = QHBoxLayout()
        lbl_assunto = QLabel("Assunto:")
        self.entry_assunto = QLineEdit()
        layout_assunto.addWidget(lbl_assunto)
        layout_assunto.addWidget(self.entry_assunto)
        layout_principal.addLayout(layout_assunto)

        # Campo de Corpo do E-mail
        layout_corpo = QVBoxLayout()
        lbl_corpo = QLabel("Corpo do E-mail:")
        self.text_corpo = QTextEdit()
        layout_corpo.addWidget(lbl_corpo)
        layout_corpo.addWidget(self.text_corpo)
        layout_principal.addLayout(layout_corpo)

        # Lista de Anexos
        layout_anexo = QVBoxLayout()
        lbl_anexo = QLabel("Anexos:")
        self.list_widget_anexos = QListWidget()
        btn_anexo = QPushButton("Selecionar Arquivos", clicked=self.selecionar_arquivos)
        layout_anexo.addWidget(lbl_anexo)
        layout_anexo.addWidget(self.list_widget_anexos)
        layout_anexo.addWidget(btn_anexo)
        layout_principal.addLayout(layout_anexo)

        # Botão de Enviar
        btn_enviar = QPushButton("Enviar E-mail", clicked=self.enviar_email)
        layout_principal.addWidget(btn_enviar)

        tab_email.setLayout(layout_principal)

    def init_config_tab(self, tab_config):
        layout_config = QVBoxLayout()

        # Configurações de E-mail
        lbl_email_config = QLabel("Configurações de E-mail")
        layout_config.addWidget(lbl_email_config)

        layout_usuario = QHBoxLayout()
        lbl_usuario = QLabel("Gmail:")
        self.entry_usuario = QLineEdit()
        layout_usuario.addWidget(lbl_usuario)
        layout_usuario.addWidget(self.entry_usuario)
        layout_config.addLayout(layout_usuario)

        layout_senha = QHBoxLayout()
        lbl_senha = QLabel("Senha:")
        self.entry_senha = QLineEdit()
        self.entry_senha.setEchoMode(QLineEdit.Password)  # Oculta a senha
        layout_senha.addWidget(lbl_senha)
        layout_senha.addWidget(self.entry_senha)
        layout_config.addLayout(layout_senha)

        # Configurações do Banco de Dados
        lbl_db_config = QLabel("Configurações do Banco de Dados")
        layout_config.addWidget(lbl_db_config)

        layout_db_host = QHBoxLayout()
        lbl_db_host = QLabel("Host:")
        self.entry_db_host = QLineEdit()
        layout_db_host.addWidget(lbl_db_host)
        layout_db_host.addWidget(self.entry_db_host)
        layout_config.addLayout(layout_db_host)

        layout_db_user = QHBoxLayout()
        lbl_db_user = QLabel("Usuário:")
        self.entry_db_user = QLineEdit()
        layout_db_user.addWidget(lbl_db_user)
        layout_db_user.addWidget(self.entry_db_user)
        layout_config.addLayout(layout_db_user)

        layout_db_password = QHBoxLayout()
        lbl_db_password = QLabel("Senha:")
        self.entry_db_password = QLineEdit()
        self.entry_db_password.setEchoMode(QLineEdit.Password)  # Oculta a senha
        layout_db_password.addWidget(lbl_db_password)
        layout_db_password.addWidget(self.entry_db_password)
        layout_config.addLayout(layout_db_password)

        layout_db_name = QHBoxLayout()
        lbl_db_name = QLabel("Nome do Banco:")
        self.entry_db_name = QLineEdit()
        layout_db_name.addWidget(lbl_db_name)
        layout_db_name.addWidget(self.entry_db_name)
        layout_config.addLayout(layout_db_name)

        # Configurações de Pastas
        lbl_pasta_config = QLabel("Configurações de Pastas")
        layout_config.addWidget(lbl_pasta_config)

        layout_pasta_origem = QHBoxLayout()
        lbl_pasta_origem = QLabel("Pasta de Origem:")
        self.entry_pasta_origem = QLineEdit()
        btn_pasta_origem = QPushButton("Selecionar", clicked=self.selecionar_pasta_origem)
        layout_pasta_origem.addWidget(lbl_pasta_origem)
        layout_pasta_origem.addWidget(self.entry_pasta_origem)
        layout_pasta_origem.addWidget(btn_pasta_origem)
        layout_config.addLayout(layout_pasta_origem)

        layout_pasta_destino = QHBoxLayout()
        lbl_pasta_destino = QLabel("Pasta de Destino:")
        self.entry_pasta_destino = QLineEdit()
        btn_pasta_destino = QPushButton("Selecionar", clicked=self.selecionar_pasta_destino)
        layout_pasta_destino.addWidget(lbl_pasta_destino)
        layout_pasta_destino.addWidget(self.entry_pasta_destino)
        layout_pasta_destino.addWidget(btn_pasta_destino)
        layout_config.addLayout(layout_pasta_destino)

        # Botão de salvar configurações
        btn_salvar = QPushButton("Salvar Configurações", clicked=self.salvar_configuracoes)
        layout_config.addWidget(btn_salvar)

        tab_config.setLayout(layout_config)

    def init_orcamentos_tab(self, tab_orcamentos):
        layout_orcamentos = QVBoxLayout()

        # Lista de orçamentos pendentes
        layout_pendentes = QVBoxLayout()
        lbl_pendentes = QLabel("Orçamentos Pendentes:")
        self.list_widget_pendentes = QListWidget()
        layout_pendentes.addWidget(lbl_pendentes)
        layout_pendentes.addWidget(self.list_widget_pendentes)
        layout_orcamentos.addLayout(layout_pendentes)

        # Lista de orçamentos enviados
        layout_enviados = QVBoxLayout()
        lbl_enviados = QLabel("Orçamentos Enviados:")
        self.list_widget_enviados = QListWidget()
        layout_enviados.addWidget(lbl_enviados)
        layout_enviados.addWidget(self.list_widget_enviados)
        layout_orcamentos.addLayout(layout_enviados)

        # Botão para enviar orçamento
        btn_enviar_orcamento = QPushButton("Enviar Orçamento", clicked=self.enviar_orcamento)
        layout_orcamentos.addWidget(btn_enviar_orcamento)

        tab_orcamentos.setLayout(layout_orcamentos)

    def preencher_email_cliente(self):
        nome_cliente = self.entry_cliente.text()
        for cliente in self.clientes:
            if cliente[0] == nome_cliente:
                self.entry_email.setText(cliente[1])
                break

    def configurar_autocompletar(self):
        nomes_clientes = [cliente[0] for cliente in self.clientes]
        completer = QCompleter(nomes_clientes)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.entry_cliente.setCompleter(completer)

    def selecionar_arquivos(self):
        arquivos, _ = QFileDialog.getOpenFileNames(
            self, "Selecione os arquivos", self.pasta_origem, "PDF Files (*.pdf);;All Files (*)"
        )
        if arquivos:
            for arquivo in arquivos:
                if arquivo not in self.caminhos_anexo:
                    self.caminhos_anexo.append(arquivo)
                    self.list_widget_anexos.addItem(arquivo)

    def enviar_email(self):
        try:
            destinatario = self.entry_email.text()
            assunto = self.entry_assunto.text()
            corpo = self.text_corpo.toPlainText()

            if not destinatario or not assunto or not corpo:
                QMessageBox.warning(self, "Aviso", "Por favor, preencha todos os campos.")
                return

            yag = yagmail.SMTP(self.gmail_usuario, self.gmail_senha)
            yag.send(
                to=destinatario,
                subject=assunto,
                contents=corpo,
                attachments=self.caminhos_anexo
            )

            QMessageBox.information(self, "Sucesso", "E-mail enviado com sucesso!")

            # Mover anexos enviados da pasta de origem para a pasta de destino
            for caminho in self.caminhos_anexo:
                if caminho.startswith(self.pasta_origem):
                    nome_arquivo = os.path.basename(caminho)
                    caminho_destino = os.path.join(self.pasta_destino, nome_arquivo)
                    try:
                        shutil.move(caminho, caminho_destino)
                    except Exception as e:
                        QMessageBox.warning(
                            self,
                            "Aviso",
                            f"Não foi possível mover o arquivo '{nome_arquivo}': {str(e)}"
                        )

            # Limpar os campos após enviar
            self.entry_email.clear()
            self.entry_assunto.clear()
            self.text_corpo.clear()
            self.list_widget_anexos.clear()
            self.caminhos_anexo.clear()

            # Atualizar as listas de orçamentos
            self.atualizar_orcamentos()

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao enviar e-mail: {str(e)}")

    def load_config(self):
        if os.path.exists(self.config_path):
            with open(self.config_path, "r") as config_file:
                config = json.load(config_file)
                self.gmail_usuario = config.get("gmail_usuario", "")
                self.gmail_senha = config.get("gmail_senha", "")
                self.db_host = config.get("db_host", "")
                self.db_usuario = config.get("db_usuario", "")
                self.db_senha = config.get("db_senha", "")
                self.db_nome = config.get("db_nome", "")
                self.pasta_origem = config.get("pasta_origem", "")
                self.pasta_destino = config.get("pasta_destino", "")

                self.entry_usuario.setText(self.gmail_usuario)
                self.entry_senha.setText(self.gmail_senha)
                self.entry_db_host.setText(self.db_host)
                self.entry_db_user.setText(self.db_usuario)
                self.entry_db_password.setText(self.db_senha)
                self.entry_db_name.setText(self.db_nome)
                self.entry_pasta_origem.setText(self.pasta_origem)
                self.entry_pasta_destino.setText(self.pasta_destino)

    def salvar_configuracoes(self):
        self.gmail_usuario = self.entry_usuario.text()
        self.gmail_senha = self.entry_senha.text()
        self.db_host = self.entry_db_host.text()
        self.db_usuario = self.entry_db_user.text()
        self.db_senha = self.entry_db_password.text()
        self.db_nome = self.entry_db_name.text()
        self.pasta_origem = self.entry_pasta_origem.text()
        self.pasta_destino = self.entry_pasta_destino.text()

        config = {
            "gmail_usuario": self.gmail_usuario,
            "gmail_senha": self.gmail_senha,
            "db_host": self.db_host,
            "db_usuario": self.db_usuario,
            "db_senha": self.db_senha,
            "db_nome": self.db_nome,
            "pasta_origem": self.pasta_origem,
            "pasta_destino": self.pasta_destino
        }

        with open(self.config_path, "w") as config_file:
            json.dump(config, config_file)

        QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")

    def selecionar_pasta_origem(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Origem")
        if pasta:
            self.entry_pasta_origem.setText(pasta)

    def selecionar_pasta_destino(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Destino")
        if pasta:
            self.entry_pasta_destino.setText(pasta)

    def load_clientes(self):
        """Carrega a lista de clientes do banco de dados."""
        try:
            cnx = mysql.connector.connect(
                host=self.db_host,
                user=self.db_usuario,
                password=self.db_senha,
                database=self.db_nome
            )
            cursor = cnx.cursor()
            cursor.execute("SELECT nome, email_adi FROM cliente")
            self.clientes = cursor.fetchall()
            cursor.close()
            cnx.close()

            # Atualiza o autocompletar com os novos clientes
            self.configurar_autocompletar()
        except mysql.connector.Error as err:
            QMessageBox.critical(self, "Erro", f"Erro ao conectar ao banco de dados: {str(err)}")

    def enviar_orcamento(self):
        orcamento_selecionado = self.list_widget_pendentes.currentItem()
        if orcamento_selecionado:
            nome_arquivo = orcamento_selecionado.text()
            caminho_origem = os.path.join(self.pasta_origem, nome_arquivo)

            # Verifica se o arquivo já está nos anexos para evitar duplicatas
            if caminho_origem not in self.caminhos_anexo:
                self.caminhos_anexo.append(caminho_origem)
                self.list_widget_anexos.addItem(caminho_origem)

            # Redireciona para a aba "Enviar E-mail"
            self.tabs.setCurrentIndex(0)  # Índice 0 corresponde à aba "Enviar E-mail"

            QMessageBox.information(
                self,
                "Orçamento Selecionado",
                f"O orçamento '{nome_arquivo}' foi adicionado aos anexos. Você pode agora selecionar o cliente e enviar o e-mail."
            )
        else:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione um orçamento pendente.")

    def atualizar_orcamentos(self):
        self.list_widget_pendentes.clear()
        self.list_widget_enviados.clear()

        if os.path.exists(self.pasta_origem):
            for arquivo in os.listdir(self.pasta_origem):
                if arquivo.endswith(".pdf"):
                    self.list_widget_pendentes.addItem(arquivo)

        if os.path.exists(self.pasta_destino):
            for arquivo in os.listdir(self.pasta_destino):
                if arquivo.endswith(".pdf"):
                    self.list_widget_enviados.addItem(arquivo)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    email_app = EmailApp()
    email_app.show()
    sys.exit(app.exec_())
