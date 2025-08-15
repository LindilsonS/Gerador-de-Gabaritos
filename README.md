# 📄 Gerador de Gabaritos e Listas de Presença  
### _Desenvolvido por Lindilson Silva_

Uma ferramenta prática e personalizável para escolas que precisam gerar **listas de presença** e **gabaritos** de forma rápida e organizada — com suporte a múltiplos formatos de arquivos e visual moderno.

---

## 🚀 Funcionalidades Principais

### ✅ Tipos de Lista
- **👩‍🏫 Lista de Alunos**: Geração de presença por **turma**.
- **🏫 Lista de Funcionários**: Geração de presença por **escola**.

### 📂 Suporte a Múltiplos Formatos CSV

#### ➤ Formato para Alunos
Campos obrigatórios:
ESCOLA | TURMA | NOME DO ALUNO | PROFESSOR REGENTE | ETAPA DE ENSINO


#### ➤ Formato para Funcionários
Campos obrigatórios:
NOME DA ESCOLA | NOME DO PROFESSOR | CPF DO PROFESSOR | ETAPA | TURMA | TURNO


### 🎨 Personalização Visual

- 10 paletas de cores elegantes:
  - Verde Suave, Rosa Delicado, Azul Sereno, Lilás Suave, Marrom Café, Cinza Elegante, Verde Menta, Roxo Real, Laranja Solar, Azul Corporativo
- Visualização prévia da cor escolhida
- Personalização do título da lista
- Campo de data opcional

### 🛠️ Recursos de Geração

- Gerar **somente listas** ou **listas + gabaritos**
- Geração de gabaritos com **1 ou 2 alunos por folha**
- **Seleção múltipla** de escolas e etapas de ensino
- Exportação em PDF (listas) e DOCX (gabaritos)

---

## 🆕 Melhorias Recentes

### 🧾 Lista de Funcionários
- Novo layout com colunas otimizadas
- Campo de assinatura ampliado (50% do espaço)
- Remoção automática de duplicatas por CPF
- Ordenação alfabética dos nomes

### 💻 Interface
- Botão seletor de tipo de lista (Alunos/Funcionários)
- Modo ajustado automaticamente ao tipo de lista
- Efeito hover em botões
- Pré-visualização de cores mais intuitiva

### 📁 Organização de Arquivos
- Estrutura de pastas por escola
- Nomes de arquivos padronizados e "limpos"
- Organização automática de saída:
lista_presenca_[turma].pdf
aluno1_gabarito.docx
aluno2_gabarito.docx
...

---

## 🧑‍💻 Como Usar

1. Inicie o aplicativo
2. Escolha o tipo de lista: **Alunos** ou **Funcionários**
3. Carregue o arquivo CSV
4. Selecione as escolas e etapas desejadas
5. Personalize o título e a data (opcional)
6. Escolha uma paleta de cores
7. Defina o modo de geração (1 ou 2 alunos por folha)
8. Clique em **GERAR DOCUMENTOS**

EXEMPLO DE MODELO WORD SELECIONADO PARA GABARITO QUE SÃO COMPATÍVEIS PARA SEREM SCANEADOS PELA PLATAFORMA ZIPGRADE: 


<img width="899" height="636" alt="image" src="https://github.com/user-attachments/assets/2eadcd8f-e92b-42ec-8ecd-1a47b338fb4c" />

PARA DOWLOAD DO MODELO: https://drive.google.com/drive/folders/1XNL_PtbTs7LvsL0QT7FnJ4BSrH5wlzBH?usp=drive_link

LINK DE ACESSO A PLATAFORMA ZIPGRADE: https://www.zipgrade.com/

---

## ⚠️ Notas Importantes

- Ao selecionar **Lista de Funcionários**, o sistema:
  - Ativa automaticamente o modo “somente lista de presença”
  - Gera uma única lista por escola
  - Remove duplicatas por CPF
  - Ajusta o layout visual para melhor legibilidade

- Saída dos arquivos:
  - **Listas**: PDF
  - **Gabaritos**: DOCX

---

## 🧪 Tecnologias Utilizadas

- 🐍 **Python**
- 🖼️ Interface gráfica: **Tkinter**
- 📝 Geração de documentos Word: **python-docx**
- 🧾 Geração de PDFs: **ReportLab**
- 📊 Processamento de dados: **Pandas**
