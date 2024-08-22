from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QListWidget, QListWidgetItem, QTabWidget, QCompleter
)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
import yagmail
import os
import sys
import json


class EmailApp(QWidget):
    def __init__(self):
        super().__init__()

        self.caminhos_anexo = []  # Lista para armazenar os caminhos dos arquivos
        self.config_path = "config.json"  # Caminho do arquivo de configuração
        self.clientes_path = "clientes.json"  # Caminho do arquivo de clientes
        self.gmail_usuario = ""  # Armazenar o Gmail do usuário
        self.gmail_senha = ""  # Armazenar a senha do Gmail do usuário
        self.clientes = []  # Lista de clientes carregados

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Envio de E-mail com Anexo")
        self.setGeometry(100, 100, 600, 600)

        # Definir ícone da janela com a imagem da logo
        self.setWindowIcon(QIcon("logo.png"))  # Substitua "logo.png" pelo nome do seu arquivo de logo

        # Configuração do TabWidget
        tabs = QTabWidget()
        tab_email = QWidget()
        tab_config = QWidget()
        tab_clientes = QWidget()

        tabs.addTab(tab_email, "Enviar E-mail")
        tabs.addTab(tab_config, "Configurações")
        tabs.addTab(tab_clientes, "Clientes")

        # Layout principal
        layout_principal = QVBoxLayout(self)
        layout_principal.addWidget(tabs)

        # Layout da aba de envio de e-mails
        self.init_email_tab(tab_email)

        # Layout da aba de configurações
        self.init_config_tab(tab_config)

        # Layout da aba de clientes
        self.init_clientes_tab(tab_clientes)

        self.setLayout(layout_principal)

        # Carregar configurações e clientes após a inicialização dos widgets
        self.load_config()
        self.load_clientes()

    def init_email_tab(self, tab_email):
        layout_principal = QVBoxLayout()

        # Campo de pesquisa de cliente
        layout_cliente = QHBoxLayout()
        lbl_cliente = QLabel("Cliente:")
        self.entry_cliente = QLineEdit()
        layout_cliente.addWidget(lbl_cliente)
        layout_cliente.addWidget(self.entry_cliente)
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

        # Configurar o auto-completar
        self.configurar_autocompletar()

    def init_config_tab(self, tab_config):
        layout_config = QVBoxLayout()

        # Campo de usuário (Gmail)
        layout_usuario = QHBoxLayout()
        lbl_usuario = QLabel("Gmail:")
        self.entry_usuario = QLineEdit()
        layout_usuario.addWidget(lbl_usuario)
        layout_usuario.addWidget(self.entry_usuario)
        layout_config.addLayout(layout_usuario)

        # Campo de senha
        layout_senha = QHBoxLayout()
        lbl_senha = QLabel("Senha:")
        self.entry_senha = QLineEdit()
        self.entry_senha.setEchoMode(QLineEdit.Password)  # Oculta a senha
        layout_senha.addWidget(lbl_senha)
        layout_senha.addWidget(self.entry_senha)
        layout_config.addLayout(layout_senha)

        # Botão de salvar configurações
        btn_salvar = QPushButton("Salvar Configurações", clicked=self.salvar_configuracoes)
        layout_config.addWidget(btn_salvar)

        tab_config.setLayout(layout_config)

    def init_clientes_tab(self, tab_clientes):
        layout_clientes = QVBoxLayout()

        # Campo de nome do cliente
        layout_nome = QHBoxLayout()
        lbl_nome = QLabel("Nome:")
        self.entry_nome = QLineEdit()
        layout_nome.addWidget(lbl_nome)
        layout_nome.addWidget(self.entry_nome)
        layout_clientes.addLayout(layout_nome)

        # Campo de e-mail do cliente
        layout_email = QHBoxLayout()
        lbl_email = QLabel("E-mail:")
        self.entry_email_cliente = QLineEdit()
        layout_email.addWidget(lbl_email)
        layout_email.addWidget(self.entry_email_cliente)
        layout_clientes.addLayout(layout_email)

        # Botão de adicionar cliente
        btn_adicionar = QPushButton("Adicionar Cliente", clicked=self.adicionar_cliente)
        layout_clientes.addWidget(btn_adicionar)

        # Lista de clientes
        self.list_widget_clientes = QListWidget()
        layout_clientes.addWidget(self.list_widget_clientes)

        # Botão de excluir cliente
        btn_excluir = QPushButton("Excluir Cliente", clicked=self.excluir_cliente)
        layout_clientes.addWidget(btn_excluir)

        # Botão de atualizar cliente
        btn_atualizar = QPushButton("Atualizar Lista", clicked=self.load_clientes)
        layout_clientes.addWidget(btn_atualizar)

        tab_clientes.setLayout(layout_clientes)

    def configurar_autocompletar(self):
        """Configurar o autocompletar para o campo de cliente."""
        nomes_clientes = [cliente['nome'] for cliente in self.clientes]
        completer = QCompleter(nomes_clientes, self)
        completer.setFilterMode(Qt.MatchContains)
        self.entry_cliente.setCompleter(completer)
        self.entry_cliente.textChanged.connect(self.preencher_email_cliente)

    def preencher_email_cliente(self):
        """Preencher o e-mail automaticamente com base na seleção do cliente."""
        nome_cliente = self.entry_cliente.text()
        for cliente in self.clientes:
            if cliente['nome'] == nome_cliente:
                self.entry_email.setText(cliente['email'])
                break
        else:
            self.entry_email.clear()

    def load_config(self):
        """Carregar configurações do arquivo JSON."""
        if os.path.exists(self.config_path):
            with open(self.config_path, "r") as file:
                config = json.load(file)
                self.gmail_usuario = config.get("gmail_usuario", "")
                self.gmail_senha = config.get("gmail_senha", "")
                self.entry_usuario.setText(self.gmail_usuario)
                self.entry_senha.setText(self.gmail_senha)

    def salvar_configuracoes(self):
        """Salvar configurações no arquivo JSON."""
        self.gmail_usuario = self.entry_usuario.text()
        self.gmail_senha = self.entry_senha.text()

        # Salvar no arquivo JSON
        with open(self.config_path, "w") as file:
            json.dump({
                "gmail_usuario": self.gmail_usuario,
                "gmail_senha": self.gmail_senha
            }, file)

        QMessageBox.information(self, "Configurações Salvas", "As configurações foram salvas com sucesso!")

    def load_clientes(self):
        """Carregar clientes do arquivo JSON e atualizar a lista."""
        if os.path.exists(self.clientes_path):
            with open(self.clientes_path, "r") as file:
                self.clientes = json.load(file)
                self.list_widget_clientes.clear()
                for cliente in self.clientes:
                    item = QListWidgetItem(f"{cliente['nome']} - {cliente['email']}")
                    self.list_widget_clientes.addItem(item)

        # Atualizar autocompletar após carregar clientes
        self.configurar_autocompletar()

    def salvar_clientes(self):
        """Salvar a lista de clientes no arquivo JSON."""
        clientes = []
        for index in range(self.list_widget_clientes.count()):
            item = self.list_widget_clientes.item(index).text()
            nome, email = item.split(' - ')
            clientes.append({"nome": nome, "email": email})

        with open(self.clientes_path, "w") as file:
            json.dump(clientes, file)

    def adicionar_cliente(self):
        """Adicionar um novo cliente à lista e salvar."""
        nome = self.entry_nome.text()
        email = self.entry_email_cliente.text()
        if not nome or not email:
            QMessageBox.critical(self, "Erro", "Por favor, preencha nome e e-mail do cliente.")
            return

        # Adicionar cliente à lista e salvar
        item = QListWidgetItem(f"{nome} - {email}")
        self.list_widget_clientes.addItem(item)
        self.salvar_clientes()

        # Limpar campos após adicionar
        self.entry_nome.clear()
        self.entry_email_cliente.clear()

    def excluir_cliente(self):
        """Excluir o cliente selecionado da lista e salvar."""
        item = self.list_widget_clientes.currentItem()
        if item:
            self.list_widget_clientes.takeItem(self.list_widget_clientes.row(item))
            self.salvar_clientes()
        else:
            QMessageBox.critical(self, "Erro", "Por favor, selecione um cliente para excluir.")

    def selecionar_arquivos(self):
        arquivos, _ = QFileDialog.getOpenFileNames(
            self, "Selecione os arquivos", "",
            "Todos os arquivos (*.*);;PDF files (*.pdf);;Excel files (*.xlsx);;Imagens (*.jpg *.jpeg *.png)"
        )

        # Adicionar arquivos à lista e exibir os nomes na interface
        for arquivo in arquivos:
            if arquivo not in self.caminhos_anexo:
                self.caminhos_anexo.append(arquivo)
                item = QListWidgetItem(os.path.basename(arquivo))  # Exibe apenas o nome do arquivo
                self.list_widget_anexos.addItem(item)

    def enviar_email(self):
        if not self.gmail_usuario or not self.gmail_senha:
            QMessageBox.critical(self, "Erro", "Por favor, configure seu Gmail e senha antes de enviar e-mails.")
            return

        destinatario = self.entry_email.text()
        assunto = self.entry_assunto.text()
        corpo = self.text_corpo.toPlainText()

        if not destinatario or not assunto or not self.caminhos_anexo:
            QMessageBox.critical(self, "Erro", "Por favor, preencha todos os campos.")
            return

        for caminho_anexo in self.caminhos_anexo:
            if not os.path.isfile(caminho_anexo):
                QMessageBox.critical(self, "Erro", f"Arquivo não encontrado: {caminho_anexo}")
                return

        try:
            yag = yagmail.SMTP(self.gmail_usuario, self.gmail_senha)
            yag.send(
                to=destinatario,
                subject=assunto,
                contents=corpo,
                attachments=self.caminhos_anexo,
            )
            QMessageBox.information(self, "Sucesso", f"E-mail enviado para {destinatario} com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao enviar o e-mail: {str(e)}")


def main():
    app = QApplication(sys.argv)
    ex = EmailApp()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
