<p align="center">
  <img src="assest/ConsulT_logo.png" alt="ConsulT Logo" width="180"/>
</p>

<h1 align="center">ConsulT</h1>

<p align="center">
  <strong>Ferramenta de auditoria e reconciliacao de clientes entre sistemas contabeis</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.10%2B-blue?logo=python&logoColor=white" alt="Python 3.10+"/>
  <img src="https://img.shields.io/badge/interface-CustomTkinter-0078D4?logo=windows&logoColor=white" alt="CustomTkinter"/>
  <img src="https://img.shields.io/badge/licen%C3%A7a-MIT-green" alt="Licenca MIT"/>
  <img src="https://img.shields.io/badge/status-em%20produ%C3%A7%C3%A3o-brightgreen" alt="Status"/>
  <img src="https://img.shields.io/badge/plataforma-Windows-lightgrey?logo=windows" alt="Windows"/>
</p>

---

## Sobre

**ConsulT** e uma aplicacao desktop desenvolvida para automatizar a verificacao e reconciliacao de cadastros de clientes entre os sistemas **e-Kontrol**, **Veri** e **e-CAC**. Elimina o trabalho manual de cruzamento de planilhas, utilizando comparacao inteligente por CNPJ/CPF e fuzzy matching de nomes.

## Funcionalidades

### Modo Veri
- Compara a base do **e-Kontrol** com a base do **e-CAC**
- Identifica empresas que constam no e-Kontrol mas **nao aparecem** no e-CAC
- Aplica 3 etapas de verificacao:
  1. Comparacao exata de CNPJ
  2. Fuzzy matching por nome completo (token sort ratio >= 90%)
  3. Fuzzy matching pelo inicio do nome (prefixo de 15 caracteres)
- Exporta relatorio `.xlsx` com as empresas nao encontradas

### Modo Consul_ECAC
- Cruzamento bidirecional entre **e-Kontrol** e uma planilha de comparacao
- Identifica:
  - **Vermelho** — empresas que sairam (estao na comparacao mas nao no e-Kontrol)
  - **Verde** — empresas a adicionar (estao no e-Kontrol mas nao na comparacao)
- Dois tipos de saida:
  - **Relatorio** — Excel com linhas coloridas (vermelho/verde)
  - **Exportacao** — Excel pronto para importacao (CNPJ formatado, contato e email)

## Screenshots

<p align="center">
  <img src="assest/ConsulT_logo.png" alt="ConsulT Interface" width="300"/>
</p>

## Tecnologias

| Tecnologia | Finalidade |
|---|---|
| [Python 3.10+](https://www.python.org/) | Linguagem principal |
| [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) | Interface grafica moderna |
| [pandas](https://pandas.pydata.org/) | Manipulacao de dados |
| [thefuzz](https://github.com/seatgeek/thefuzz) | Fuzzy matching de strings |
| [openpyxl](https://openpyxl.readthedocs.io/) | Leitura/escrita de Excel com estilos |
| [Pillow](https://python-pillow.org/) | Carregamento de imagens (logo) |
| [Unidecode](https://pypi.org/project/Unidecode/) | Normalizacao de caracteres acentuados |

## Instalacao

### Pre-requisitos

- Python 3.10 ou superior
- pip (gerenciador de pacotes)

### Passos

```bash
# Clone o repositorio
git clone https://github.com/seu-usuario/ConsulT.git
cd ConsulT

# Instale as dependencias
pip install pandas openpyxl customtkinter thefuzz python-Levenshtein unidecode Pillow

# Execute
python ConsulT.py
```

## Como Usar

1. **Selecione o modo** — `Veri` ou `Consul_ECAC`
2. **Carregue os arquivos** — selecione a planilha do e-Kontrol e a planilha de comparacao (e-CAC ou outra)
3. **Defina a saida** — escolha o nome e local do arquivo de resultado
4. **(Consul_ECAC)** Escolha o tipo de saida: Relatorio com cores ou Exportacao
5. **Clique em Executar** — acompanhe o progresso no log em tempo real

## Estrutura do Projeto

```
ConsulT/
├── ConsulT.py              # Aplicacao principal (logica + interface)
├── README.md               # Documentacao
└── assest/
    ├── ConsulT_logo.png    # Logo do projeto
    ├── favicon.ico         # Icone da janela
    ├── favicon-16x16.png   # Favicon 16px
    ├── favicon-32x32.png   # Favicon 32px
    └── ...                 # Outros assets
```

## Autor

<table>
  <tr>
    <td align="center">
      <strong>Hugo L. Almeida</strong><br/>
      Desenvolvedor RPA & Automacao
    </td>
  </tr>
</table>

## Licenca

Este projeto esta licenciado sob a [Licenca MIT](LICENSE).

---

<p align="center">
  Feito com dedicacao por <strong>Hugo L. Almeida</strong>
</p>
