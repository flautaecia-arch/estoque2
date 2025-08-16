# Manual de Uso - Sistema de Contagem de Estoque

## 🎯 Guia Passo a Passo

### 1. Primeiro Acesso

1. **Inicie o aplicativo**
   ```bash
   cd estoque_app
   source venv/bin/activate
   python src/main.py
   ```

2. **Abra no navegador**
   - Acesse: `http://localhost:5000`
   - Você verá 3 abas: Produtos, Contagem, Relatórios

### 2. Importando Produtos (Recomendado)

#### Preparar Planilha Excel
- **Formato**: .xlsx ou .xls
- **Estrutura**: 
  - Coluna A: Códigos dos produtos (números ou texto)
  - Coluna B: Nomes dos produtos
- **Exemplo**:
  ```
  1206    VODKA DIAMONDS
  1213    VODKA EXTREME
  1329    MISS ELYSEES
  ```

#### Processo de Importação
1. Clique na aba **"Produtos"**
2. Role até a seção **"Importar Produtos via Excel"**
3. Clique em **"Selecionar Arquivo Excel"**
4. Escolha sua planilha
5. Clique em **"Importar Produtos"**
6. Aguarde o processamento
7. Veja o resultado da importação

#### Resultado da Importação
O sistema mostrará:
- ✅ **Produtos importados**: Novos produtos adicionados
- ⚠️ **Produtos duplicados**: Códigos que já existiam
- ❌ **Produtos com erro**: Linhas com problemas
- 📊 **Total processado**: Total de linhas da planilha

### 3. Cadastro Manual de Produtos (Alternativo)

Se preferir cadastrar um por vez:

1. Na aba **"Produtos"**
2. Preencha:
   - **Código do Produto**: Ex: 1206
   - **Nome do Produto**: Ex: VODKA DIAMONDS
3. Clique **"Salvar Produto"**

### 4. Fazendo Contagem de Estoque

#### Buscar Produto
1. Vá para a aba **"Contagem"**
2. Digite o código do produto no campo **"Código do Produto"**
3. Clique **"Buscar"**
4. O produto aparecerá se existir

#### Registrar Lote
Após encontrar o produto:
1. Preencha os dados do lote:
   - **Lote**: Ex: L001, LOTE123, etc.
   - **Mês de Validade**: Selecione o mês
   - **Ano de Validade**: Ex: 2025, 2026
   - **Quantidade**: Ex: 50, 100

2. Clique **"Registrar Contagem"**

#### ⚠️ Importante: Consolidação Automática
- Se você registrar o **mesmo lote** novamente, o sistema **somará** as quantidades
- Exemplo:
  - 1ª contagem: Lote L001 = 50 unidades
  - 2ª contagem: Lote L001 = 30 unidades
  - **Resultado**: Lote L001 = 80 unidades (50+30)

### 5. Visualizando Contagens

Na aba **"Contagem"**, role até **"Contagens Registradas"**:
- Veja todas as contagens feitas
- Informações: Código, Produto, Lote, Validade, Quantidade
- Pode excluir contagens se necessário

### 6. Gerando Relatórios

#### Resumo em Tempo Real
Na aba **"Relatórios"**:
- Veja o **"Resumo do Estoque"**
- Mostra total por produto
- Total geral de todos os produtos
- Clique **"Atualizar"** para refresh

#### Relatório PDF
1. Clique **"Gerar PDF"**
2. Arquivo será baixado automaticamente
3. **Conteúdo**:
   - Resumo por produto (ordenado por código)
   - Detalhamento por lote
   - Totais e subtotais

#### Relatório Excel
1. Clique **"Gerar Excel"**
2. Arquivo será baixado automaticamente
3. **Conteúdo**:
   - **Aba 1**: Resumo por produto
   - **Aba 2**: Detalhado por lote
   - Formatação profissional

## 🔄 Fluxo de Trabalho Recomendado

### Para Contagem Completa do Estoque:

1. **Preparação**
   - Organize planilha com todos os produtos
   - Importe via Excel (uma única vez)

2. **Contagem Física**
   - Vá ao estoque com tablet/celular
   - Acesse o sistema pelo navegador
   - Para cada produto encontrado:
     - Busque pelo código
     - Registre lote, validade e quantidade

3. **Consolidação**
   - Se encontrar mais do mesmo lote, registre novamente
   - O sistema somará automaticamente

4. **Relatório Final**
   - Gere PDF ou Excel
   - Tenha relatório completo com totais

### Para Contagem Parcial/Específica:

1. **Cadastro Rápido**
   - Cadastre apenas os produtos que vai contar
   - Use cadastro manual ou mini-planilha

2. **Contagem Direcionada**
   - Foque nos produtos específicos
   - Registre todas as informações

3. **Relatório Parcial**
   - Gere relatório dos produtos contados

## 🚨 Dicas Importantes

### ✅ Boas Práticas
- **Códigos únicos**: Cada produto deve ter código único
- **Lotes consistentes**: Use padrão para nomear lotes
- **Backup regular**: Faça backup do arquivo `app.db`
- **Teste primeiro**: Teste com poucos produtos antes da contagem completa

### ⚠️ Cuidados
- **Não feche o navegador** durante contagem longa
- **Confirme dados** antes de registrar
- **Internet estável** se usar em dispositivo móvel
- **Planilha limpa** para importação (sem células vazias)

### 🔧 Solução de Problemas

#### Produto não encontrado
- Verifique se o código está correto
- Confirme se foi importado/cadastrado
- Códigos são case-sensitive

#### Erro na importação
- Verifique formato da planilha (.xlsx ou .xls)
- Confirme se há pelo menos 2 colunas
- Remova linhas vazias da planilha

#### Relatório não gera
- Verifique se há contagens registradas
- Aguarde alguns segundos para processamento
- Tente atualizar a página

#### Sistema lento
- Feche outras abas do navegador
- Verifique conexão de internet
- Reinicie o aplicativo se necessário

## 📱 Uso em Dispositivos Móveis

O sistema é totalmente responsivo:

### Smartphone/Tablet
- Interface se adapta automaticamente
- Botões maiores para toque
- Tabelas com scroll horizontal
- Formulários otimizados

### Dicas Mobile
- Use modo paisagem para tabelas
- Zoom para campos pequenos se necessário
- Mantenha conexão estável
- Considere usar offline (dados ficam no dispositivo)

## 🎯 Casos de Uso Comuns

### Loja/Varejo
- Importe catálogo de produtos
- Conte estoque por seção
- Gere relatório para reposição

### Farmácia
- Foque nas datas de validade
- Use lotes para controle sanitário
- Relatório para órgãos fiscalizadores

### Indústria
- Conte matéria-prima por lote
- Controle produtos acabados
- Relatório para produção

### Distribuidora
- Importe grande volume de produtos
- Conte por fornecedor/marca
- Relatório para vendas

---

**💡 Lembre-se: O sistema foi feito para facilitar sua vida. Explore as funcionalidades e adapte ao seu processo de trabalho!**

