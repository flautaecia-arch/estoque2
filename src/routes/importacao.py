from flask import Blueprint, request, jsonify
from src.models.user import db
from src.models.produto import Produto
from sqlalchemy.exc import IntegrityError
import pandas as pd
import io

importacao_bp = Blueprint('importacao', __name__)

@importacao_bp.route('/importar/produtos', methods=['POST'])
def importar_produtos():
    """Importa produtos de uma planilha Excel"""
    try:
        # Verificar se foi enviado um arquivo
        if 'arquivo' not in request.files:
            return jsonify({'erro': 'Nenhum arquivo foi enviado'}), 400
        
        arquivo = request.files['arquivo']
        
        if arquivo.filename == '':
            return jsonify({'erro': 'Nenhum arquivo foi selecionado'}), 400
        
        # Verificar se é um arquivo Excel
        if not arquivo.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'erro': 'Arquivo deve ser uma planilha Excel (.xlsx ou .xls)'}), 400
        
        # Ler o arquivo Excel
        try:
            # Tentar ler como bytes primeiro
            arquivo_bytes = arquivo.read()
            df = pd.read_excel(io.BytesIO(arquivo_bytes))
        except Exception as e:
            return jsonify({'erro': f'Erro ao ler arquivo Excel: {str(e)}'}), 400
        
        # Verificar se o DataFrame não está vazio
        if df.empty:
            return jsonify({'erro': 'Planilha está vazia'}), 400
        
        # Verificar se há pelo menos 2 colunas
        if len(df.columns) < 2:
            return jsonify({'erro': 'Planilha deve ter pelo menos 2 colunas (código e nome)'}), 400
        
        # Normalizar nomes das colunas (remover espaços e converter para minúsculas)
        df.columns = df.columns.str.strip().str.lower()
        
        # Verificar se as colunas necessárias existem
        colunas_possiveis_codigo = ['codigo', 'código', 'code', 'cod']
        colunas_possiveis_nome = ['nome', 'produto', 'name', 'descricao', 'descrição', 'description']
        
        coluna_codigo = None
        coluna_nome = None
        
        # Primeiro, tentar encontrar colunas com nomes conhecidos
        for col in df.columns:
            if col in colunas_possiveis_codigo:
                coluna_codigo = col
            if col in colunas_possiveis_nome:
                coluna_nome = col
        
        # Se não encontrou colunas com nomes conhecidos, usar as duas primeiras colunas
        if not coluna_codigo or not coluna_nome:
            if len(df.columns) >= 2:
                coluna_codigo = df.columns[0]  # Primeira coluna como código
                coluna_nome = df.columns[1]    # Segunda coluna como nome
                print(f"Usando primeira coluna '{coluna_codigo}' como código e segunda coluna '{coluna_nome}' como nome")
            else:
                return jsonify({
                    'erro': 'Não foi possível identificar as colunas de código e nome. Certifique-se de que a planilha tem pelo menos 2 colunas.'
                }), 400
        
        # Processar os dados
        produtos_importados = 0
        produtos_duplicados = 0
        produtos_erro = 0
        erros_detalhados = []
        
        for index, row in df.iterrows():
            try:
                # Verificar se os valores não são nulos ou vazios
                codigo = str(row[coluna_codigo]).strip() if pd.notna(row[coluna_codigo]) else ''
                nome = str(row[coluna_nome]).strip() if pd.notna(row[coluna_nome]) else ''
                
                if not codigo or not nome or codigo.lower() == 'nan' or nome.lower() == 'nan':
                    produtos_erro += 1
                    erros_detalhados.append(f'Linha {index + 2}: Código ou nome vazio/inválido')
                    continue
                
                # Verificar se o produto já existe
                produto_existente = Produto.query.filter_by(codigo=codigo).first()
                if produto_existente:
                    produtos_duplicados += 1
                    continue
                
                # Criar novo produto
                novo_produto = Produto(codigo=codigo, nome=nome)
                db.session.add(novo_produto)
                produtos_importados += 1
                
            except Exception as e:
                produtos_erro += 1
                erros_detalhados.append(f'Linha {index + 2}: {str(e)}')
        
        # Salvar no banco de dados
        try:
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            return jsonify({'erro': f'Erro ao salvar no banco de dados: {str(e)}'}), 500
        
        # Preparar resposta
        resultado = {
            'mensagem': 'Importação concluída',
            'produtos_importados': produtos_importados,
            'produtos_duplicados': produtos_duplicados,
            'produtos_erro': produtos_erro,
            'total_processados': len(df)
        }
        
        if erros_detalhados:
            resultado['erros_detalhados'] = erros_detalhados[:10]  # Limitar a 10 erros para não sobrecarregar
            if len(erros_detalhados) > 10:
                resultado['erros_detalhados'].append(f'... e mais {len(erros_detalhados) - 10} erros')
        
        return jsonify(resultado), 200
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'erro': f'Erro interno: {str(e)}'}), 500

@importacao_bp.route('/template/produtos', methods=['GET'])
def baixar_template():
    """Gera um template Excel para importação de produtos"""
    try:
        # Criar DataFrame com template
        template_data = {
            'codigo': ['PROD001', 'PROD002', 'PROD003'],
            'nome': ['Produto Exemplo 1', 'Produto Exemplo 2', 'Produto Exemplo 3']
        }
        
        df = pd.DataFrame(template_data)
        
        # Criar arquivo Excel em memória
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Produtos', index=False)
            
            # Ajustar largura das colunas
            worksheet = writer.sheets['Produtos']
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 40
        
        buffer.seek(0)
        
        from flask import send_file
        return send_file(
            buffer,
            as_attachment=True,
            download_name='template_produtos.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'erro': f'Erro ao gerar template: {str(e)}'}), 500

