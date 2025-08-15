# Gerador de Gabaritos e Listas de Presença

Um aplicativo para gerar gabaritos e listas de presença para escolas, com suporte a múltiplos formatos e tipos de lista.

## Funcionalidades Principais

### 1. Dois Tipos de Lista
- **Lista de Alunos**: Para gerar listas de presença de alunos por turma
- **Lista de Funcionários**: Para gerar listas de presença de funcionários por escola

### 2. Suporte a Múltiplos Formatos CSV
O sistema aceita dois formatos de arquivo CSV:

#### Formato de Alunos
Colunas necessárias:
- ESCOLA
- TURMA
- NOME DO ALUNO
- PROFESSOR REGENTE
- ETAPA DE ENSINO

#### Formato de Funcionários
Colunas necessárias:
- NOME DA ESCOLA ou ESCOLA
- NOME DO PROFESSOR
- CPF DO PROFESSOR
- ETAPA
- TURMA
- TURNO

### 3. Personalização Visual
- **10 Paletas de Cores**: Verde Suave, Rosa Delicado, Azul Sereno, Lilás Suave, Marrom Café, Cinza Elegante, Verde Menta, Roxo Real, Laranja Solar, Azul Corporativo
- Visualização prévia das cores selecionadas
- Personalização do título da lista
- Campo para data opcional

### 4. Recursos de Geração
- Opção de gerar apenas lista de presença
- Opção de gerar lista de presença junto com gabaritos
- Modo de geração com um ou dois alunos por folha (para gabaritos)
- Seleção múltipla de escolas
- Seleção múltipla de etapas de ensino

## Melhorias Recentes

### 1. Lista de Funcionários
- Novo formato otimizado com colunas ajustadas
- Campo de assinatura mais amplo
- Distribuição equilibrada entre nome e assinatura (50/50)
- Remoção automática de duplicatas baseada no CPF
- Ordenação alfabética automática

### 2. Interface
- Novo botão de tipo de lista
- Ajuste automático de modo quando "Lista de Funcionários" é selecionada
- Botão de geração com efeito hover
- Preview de cores mais intuitivo

### 3. Organização de Arquivos
- Estrutura de pastas otimizada por escola
- Nomenclatura padronizada dos arquivos
- Sanitização automática de nomes de arquivo

## Como Usar

1. Inicie o aplicativo
2. Selecione o tipo de lista (Alunos ou Funcionários)
3. Carregue o arquivo CSV com os dados
4. Selecione a(s) escola(s) desejada(s)
5. Escolha a etapa de ensino (se aplicável)
6. Personalize o título e a data (opcional)
7. Escolha uma paleta de cores
8. Selecione o modo de geração (para gabaritos)
9. Clique em "GERAR DOCUMENTOS"

## Notas Importantes

- Ao selecionar "Lista de Funcionários", o sistema automaticamente:
  - Força o modo de lista única
  - Ativa "apenas lista de presença"
  - Remove duplicatas de funcionários
  - Ajusta o layout para melhor visualização
- O sistema salvará os arquivos em uma estrutura organizada por escola
- As listas de presença são geradas em PDF
- Os gabaritos são gerados em DOCX
│       ├── lista_presenca_[turma].pdf
│       ├── aluno1_gabarito.docx
│       ├── aluno2_gabarito.docx
│       └── ...
```

## Observações Técnicas
- Desenvolvido em Python
- Interface gráfica com Tkinter
- Geração de PDF com ReportLab
- Manipulação de DOCX com python-docx
- Processamento de dados com Pandas
