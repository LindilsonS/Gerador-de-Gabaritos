# 📚 Gerador de Gabaritos e Lista de Presença

Sistema desenvolvido para auxiliar escolas na geração automatizada de gabaritos personalizados e listas de presença.

## 🚀 Funcionalidades

- ✨ Geração automática de gabaritos individuais por aluno
- 📋 Criação de listas de presença organizadas por turma
- 🔄 Substituição dinâmica de variáveis em documentos Word
- 📁 Organização automática em pastas por escola/turma
- 🎨 Interface gráfica intuitiva
- 🔍 Suporte a caracteres especiais nos nomes

## 💻 Requisitos do Sistema

- Python 3.7 ou superior

## 📥 Como Instalar

1. Instale as dependências:

   pip install -r requirements.txt

## 🎯 Como Usar

1. Execute o programa:

   python main.py

2. Na interface gráfica:
   - Selecione o arquivo CSV com dados dos alunos
   - Escolha o modelo Word (.docx) para gabaritos
   - Selecione a pasta onde serão salvos os arquivos
   - Clique em "Gerar Documentos"

### 📊 Estrutura do CSV

O arquivo CSV deve estar formatado com ponto e vírgula (;) e conter as colunas:

| ESCOLA | TURMA | NOME DO ALUNO | PROFESSOR REGENTE |
|--------|-------|---------------|-------------------|
| Escola Municipal | 5º Ano A | João Silva | Maria Santos |

### 📝 Modelo Word

No documento modelo (.docx), utilize os seguintes marcadores:
- `$VARIÁVEL NOME DO ALUNO`
- `$VARIÁVEL ESCOLA`
- `$VARIÁVEL TURMA`
- `$VARIÁVEL PROFESSOR REGENTE`

### 📂 Organização dos Arquivos Gerados

```
Pasta Selecionada/
├── Escola Municipal/
│   ├── 5º Ano A/
│   │   ├── joao_silva_gabarito.docx
│   │   ├── maria_oliveira_gabarito.docx
│   │   └── lista_presenca_5_Ano_A.pdf
│   └── 5º Ano B/
└── Outra Escola/
```

## 🛠️ Recursos Técnicos

- ✅ Sanitização automática de nomes de arquivos
- 📄 Suporte para substituições em:
  - Parágrafos
  - Tabelas
  - Caixas de texto
- 📊 Lista de presença com design profissional
- 🔢 Numeração automática de alunos
- ⚠️ Tratamento robusto de erros

## 📋 Lista de Presença

As listas de presença incluem:
- Cabeçalho com dados da escola e turma
- Numeração automática dos alunos
- Campo para data
- Contagem total de alunos
- Campo para assinaturas
- Rodapé institucional

## ⚠️ Observações Importantes

1. Verifique se os marcadores no modelo Word estão corretos
2. O CSV deve estar codificado em UTF-8
3. Nomes com caracteres especiais são tratados automaticamente


## 👥 Autor

- Lindilson Silva
