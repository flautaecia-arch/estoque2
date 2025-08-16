# Sistema de Contagem de Estoque

Um aplicativo web completo para gerenciamento e contagem de estoque com funcionalidades avançadas de importação e relatórios.

## 🚀 Funcionalidades

### ✅ Gestão de Produtos
- Cadastro manual de produtos (código e nome)
- **Importação em massa via planilha Excel**
- Listagem e exclusão de produtos
- Busca de produtos por código

### ✅ Contagem de Estoque
- Busca rápida de produtos por código
- Registro de lotes com:
  - Número do lote
  - Mês e ano de validade
  - Quantidade disponível
- **Consolidação automática de lotes duplicados** (soma quantidades)
- Visualização de todas as contagens registradas

### ✅ Relatórios Avançados
- **Relatório em PDF** com formatação profissional
- **Relatório em Excel** com múltiplas abas
- Ordenação automática por código do produto
- Subtotais por produto e total geral
- Resumo em tempo real na interface

### ✅ Interface Moderna
- Design responsivo (funciona em desktop e mobile)
- Interface intuitiva com abas organizadas
- Feedback visual para todas as operações
- Animações e transições suaves

## 📋 Requisitos

- Python 3.11+
- Flask
- SQLAlchemy
- Pandas (para importação Excel)
- ReportLab (para PDF)
- OpenPyXL (para Excel)

## 🛠️ Instalação

1. **Clone ou baixe o projeto**
```bash
cd estoque_app
```

2. **Ative o ambiente virtual**
```bash
source venv/bin/activate
```

3. **Instale as dependências**
```bash
pip install -r requirements.txt
```

4. **Execute o aplicativo**
```bash
python src/main.py
```

5. **Acesse no navegador**
```
http://localhost:5000
```

## 📊 Como Usar

### 1. Importar Produtos via Excel

1. Prepare uma planilha Excel (.xlsx ou .xls) com:
   - **Primeira coluna**: Códigos dos produtos
   - **Segunda coluna**: Nomes dos produtos
   
2. Na aba "Produtos", clique em "Importar Produtos"
3. Selecione sua planilha e clique "Importar Produtos"
4. O sistema processará automaticamente todos os produtos

**Exemplo de planilha:**
```
| 1206 | VODKA DIAMONDS    |
| 1213 | VODKA EXTREME     |
| 1329 | MISS ELYSEES      |
```

### 2. Fazer Contagem de Estoque

1. Vá para a aba "Contagem"
2. Digite o código do produto e clique "Buscar"
3. Preencha os dados do lote:
   - Número do lote
   - Mês de validade
   - Ano de validade
   - Quantidade
4. Clique "Registrar Contagem"

**Importante**: Se você registrar o mesmo lote novamente, o sistema somará as quantidades automaticamente!

### 3. Gerar Relatórios

1. Vá para a aba "Relatórios"
2. Clique em "Gerar PDF" ou "Gerar Excel"
3. O arquivo será baixado automaticamente

**Os relatórios incluem:**
- Lista ordenada por código do produto
- Subtotal de cada produto
- Total geral de todos os produtos
- Detalhamento por lote (no Excel)

## 🏗️ Estrutura do Projeto

```
estoque_app/
├── src/
│   ├── models/          # Modelos do banco de dados
│   │   ├── produto.py   # Modelo de produtos
│   │   ├── contagem.py  # Modelo de contagens
│   │   └── user.py      # Configuração do banco
│   ├── routes/          # Rotas da API
│   │   ├── produto.py   # CRUD de produtos
│   │   ├── contagem.py  # Gestão de contagens
│   │   ├── relatorio.py # Geração de relatórios
│   │   └── importacao.py # Importação Excel
│   ├── static/          # Frontend
│   │   ├── index.html   # Interface principal
│   │   ├── styles.css   # Estilos
│   │   └── script.js    # JavaScript
│   ├── database/        # Banco de dados SQLite
│   └── main.py          # Aplicação principal
├── venv/                # Ambiente virtual
├── requirements.txt     # Dependências
└── README.md           # Esta documentação
```

## 🔧 API Endpoints

### Produtos
- `GET /api/produtos` - Listar produtos
- `POST /api/produtos` - Criar produto
- `GET /api/produtos/<codigo>` - Buscar por código
- `DELETE /api/produtos/<id>` - Excluir produto

### Contagens
- `GET /api/contagens` - Listar contagens
- `POST /api/contagens` - Registrar contagem
- `GET /api/contagens/produto/<codigo>` - Contagens de um produto
- `DELETE /api/contagens/<id>` - Excluir contagem

### Relatórios
- `GET /api/relatorio/resumo` - Resumo JSON
- `GET /api/relatorio/pdf` - Baixar PDF
- `GET /api/relatorio/excel` - Baixar Excel

### Importação
- `POST /api/importar/produtos` - Importar Excel
- `GET /api/template/produtos` - Baixar template

## 🎯 Características Técnicas

### Backend (Flask)
- Arquitetura modular com blueprints
- Banco de dados SQLite com SQLAlchemy
- CORS habilitado para frontend
- Validação de dados robusta
- Tratamento de erros abrangente

### Frontend (HTML/CSS/JS)
- Interface responsiva com CSS Grid/Flexbox
- JavaScript vanilla para máxima compatibilidade
- Fetch API para comunicação com backend
- Feedback visual em tempo real
- Design moderno com gradientes e animações

### Banco de Dados
- **Produtos**: id, codigo (único), nome
- **Contagens**: id, produto_id, lote, validade_mes, validade_ano, quantidade
- Constraint único para produto+lote (evita duplicação)
- Relacionamento com cascade delete

## 🚀 Funcionalidades Avançadas

### Consolidação de Lotes
O sistema automaticamente:
- Detecta lotes duplicados (mesmo produto + mesmo lote)
- Soma as quantidades ao invés de criar registros duplicados
- Atualiza a validade com os dados mais recentes

### Importação Inteligente
- Detecta automaticamente colunas de código e nome
- Suporta variações de nomes (codigo, código, code, etc.)
- Usa primeira e segunda colunas se não encontrar nomes padrão
- Relatório detalhado de importação com estatísticas

### Relatórios Profissionais
- **PDF**: Formatação profissional com tabelas e totais
- **Excel**: Múltiplas abas (resumo + detalhado)
- Ordenação automática por código
- Cálculos automáticos de subtotais

## 🔒 Segurança

- Validação de tipos de arquivo (apenas .xlsx/.xls)
- Sanitização de dados de entrada
- Proteção contra SQL injection (SQLAlchemy ORM)
- Tratamento seguro de uploads de arquivo

## 📱 Compatibilidade

- **Navegadores**: Chrome, Firefox, Safari, Edge
- **Dispositivos**: Desktop, tablet, mobile
- **Sistemas**: Windows, macOS, Linux
- **Planilhas**: Excel (.xlsx), Excel 97-2003 (.xls)

## 🎨 Personalização

O sistema foi desenvolvido com foco na usabilidade e pode ser facilmente personalizado:

- **Cores**: Modifique as variáveis CSS em `styles.css`
- **Layout**: Ajuste a estrutura em `index.html`
- **Funcionalidades**: Adicione novos endpoints em `routes/`
- **Campos**: Estenda os modelos em `models/`

## 📞 Suporte

Para dúvidas ou problemas:
1. Verifique se todas as dependências estão instaladas
2. Confirme que o banco de dados foi criado corretamente
3. Verifique os logs do Flask para erros específicos
4. Teste com planilhas de exemplo primeiro

---

**Desenvolvido com ❤️ para facilitar o controle de estoque**

