# 🏥 Gerador de QR Codes - Inventário de Equipamentos

Ferramenta web para gerar automaticamente QR Codes de equipamentos a partir de uma planilha Excel. As etiquetas geradas incluem o QR Code e a TAG do equipamento visível abaixo, facilitando a identificação e evitando erros na hora de colar no equipamento.

## 🎯 Objetivo

Automatizar a criação de etiquetas com QR Code para inventário de equipamentos, permitindo que o processo de contagem seja feito por leitura de código, de forma mais rápida e confiável.

## ✨ Funcionalidades

- Upload de planilha `.xlsx` com a lista de equipamentos
- Suporte a planilhas com múltiplas abas
- Seleção da coluna que contém a TAG (padrão: `TAG`)
- Geração de QR Codes em lote com a TAG impressa abaixo do código
- Download do arquivo Excel com os QR Codes já embutidos nas células
- Processamento em memória (sem armazenar dados no servidor)
- Interface simples, sem necessidade de instalação

## 📋 Formato da planilha

A planilha deve conter, no mínimo, uma coluna com a TAG ou identificador único do equipamento. Exemplo:

| TAG | Equipamento | Modelo | Fabricante | Setor | Nº Série |
|---|---|---|---|---|---|
| TAG-02332 | MEDIDOR DE SINAIS | CARESCAPE | GE HEALTHCARE | CENTRO CIRÚRGICO | 123456ABC |

A coluna padrão para geração do QR Code é a **TAG**, mas outra coluna pode ser selecionada na interface, caso necessário.

## 🚀 Como usar

### Versão online (Streamlit Cloud)

Acesse aqui: [![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://gerador-qr-inventario.streamlit.app/)

1. Faça o upload da planilha `.xlsx`
2. Selecione a aba (se houver mais de uma)
3. Confirme a coluna que contém a TAG
4. Clique em **Gerar QR Codes**
5. Baixe a planilha com os QR Codes gerados

### Versão local (para desenvolvimento)

```bash
# Clone o repositório
git clone https://github.com/Mir4nd4s/gerador-qr-inventario.git
cd gerador-qr-inventario

# Crie um ambiente virtual
python -m venv .venv

# Ative o ambiente
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

# Instale as dependências
pip install -r requirements.txt

# Rode a aplicação
streamlit run app.py
```

A aplicação abrirá automaticamente em http://localhost:8501.

🛠️ Tecnologias utilizadas

    Python 3.10+ - Linguagem base

    Streamlit - Interface web

    pandas - Leitura de planilhas

    openpyxl - Manipulação de arquivos Excel

    qrcode - Geração dos QR Codes

    Pillow - Manipulação de imagens

📁 Estrutura do projeto

```text

gerador-qr-inventario/
├── app.py               # Aplicação principal
├── requirements.txt     # Dependências do projeto
└── README.md            # Este arquivo
```

🔒 Privacidade e segurança

    Os arquivos enviados são processados em memória e não são armazenados em disco

    Nenhum dado é enviado para terceiros

    O código-fonte fica no GitHub, mas os dados das planilhas nunca saem da sessão do usuário

🤝 Contribuições

Sugestões de melhoria são bem-vindas. Abra uma issue ou envie um pull request.
📄 Licença

## 🔗 Projeto relacionado

- **[PWA Inventário de Equipamentos](https://github.com/Mir4nd4s/pwa-inventario)** — Aplicativo mobile usado pelos técnicos para escanear os QR Codes gerados por esta ferramenta durante o inventário de campo. Funciona offline no celular.

> ⚠️ **Nota:** o repositório da PWA é privado por conter informações operacionais internas (lista de unidades e setores). O acesso é restrito à equipe responsável pelo inventário.

Este projeto é de uso interno. Todos os direitos reservados.
👤 Autor

Mir4nd4s

    GitHub: @Mir4nd4s
