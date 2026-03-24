# 📊 Comparador de Dados (Console)

Aplicação desenvolvida em C# para comparação de dados entre arquivos Excel (.xlsx), CSV e TXT, com foco em validação, auditoria e reconciliação de informações.

## 🔥 Funcionalidades

- 📂 Leitura automática de arquivos:
  - Excel (.xlsx)
  - CSV (.csv)
  - TXT (.txt)

- 🧠 Detecção automática:
  - Tipo de arquivo
  - Separador (CSV/TXT)
  - Estrutura de colunas

- 🔍 Comparação dinâmica:
  - Definição de coluna identificadora (ID)
  - Comparação de múltiplas colunas
  - Tratamento de dados inconsistentes

- ⚡ Validações robustas:
  - Colunas inexistentes
  - Valores nulos ou vazios
  - IDs duplicados
  - Diferenças de formatação (case, espaços, BOM)

- 📊 Geração de relatório em Excel:
  - Aba **Apenas_A** (registros só na base A)
  - Aba **Apenas_B** (registros só na base B)
  - Aba **Divergencias** (diferenças entre os dados)

- 📈 Visão geral da comparação:
  - Registros exclusivos por base
  - Registros em comum
  - Registros com divergência

- 🚀 Recursos extras:
  - Abertura automática do Excel após geração
  - Validação de arquivo em uso (evita sobrescrita)
  - Código preparado para grandes volumes de dados

---

## 🧠 Tecnologias utilizadas

- C#
- .NET
- EPPlus (manipulação de Excel)
- LINQ

---

## 🎯 Objetivo

Criar uma ferramenta genérica e reutilizável para análise e comparação de dados, simulando cenários reais de mercado como:

- Reconciliação de bases
- Validação de cargas
- Auditoria de dados
- Integração entre sistemas

---

## 🚀 Próximos passos

- Exportação com gráficos e dashboards

---

## 👨‍💻 Autor

Desenvolvido por Miguel de Souza
