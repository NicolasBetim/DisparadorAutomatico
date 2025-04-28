
# Disparador Automático de Mensagens WhatsApp 📲

Este projeto é um **sistema automático de envio de mensagens no WhatsApp** que utiliza **planilhas Excel (.xlsx)** para armazenar os contatos e envia mensagens personalizadas (com ou sem anexos) de forma simples e prática via **WhatsApp Web**.

A interface gráfica foi construída com **Tkinter**, e o navegador é controlado usando **Selenium WebDriver**.

---

## ✨ Funcionalidades

- Selecionar uma **planilha Excel** com os números de telefone.
- Escrever uma **mensagem personalizada**.
- Opcionalmente, **anexar arquivos** (imagens, documentos, vídeos).
- Controle de envio: exibe o **progresso** e o **status** durante o envio.
- Sistema de **re-tentativas automáticas** em caso de falha no envio.

---

## 🔧 Tecnologias Utilizadas

- Python 3
- Tkinter
- Selenium WebDriver
- WebDriver Manager
- Openpyxl
- Pyperclip

---

## 📋 Requisitos

Antes de rodar o projeto, você precisa ter:

- **Python 3.8+** instalado ([download aqui](https://www.python.org/downloads/))
- Google Chrome atualizado
- Internet estável para acesso ao WhatsApp Web
- Instalar as dependências:

```bash
pip install selenium webdriver-manager openpyxl pyperclip
```

---

## 🚀 Como usar

1. Clone este repositório ou baixe o projeto.
2. Execute o script principal.
3. Escolha sua planilha (.xlsx).
4. Digite sua mensagem e selecione anexos (opcional).
5. Clique em **Iniciar envio**.

---

## 📝 Observações

- A planilha deve conter uma coluna chamada `Telefone` com os números no formato internacional (ex: 5511999999999).
- O envio é feito via **WhatsApp Web**, então será necessário escanear o QR Code na primeira vez.
- Certifique-se de que o navegador Chrome esteja instalado.

---

## 📄 Licença

Este projeto é de uso livre para fins educacionais.
