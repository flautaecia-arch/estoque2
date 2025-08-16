from flask import Blueprint, request, jsonify
from src.models.user import db
from src.models.produto import Produto
from src.models.contagem import Contagem
from sqlalchemy.exc import IntegrityError
from sqlalchemy import func

contagem_bp = Blueprint('contagem', __name__)

@contagem_bp.route('/contagens', methods=['GET'])
def listar_contagens():
    """Lista todas as contagens com informações do produto"""
    contagens = db.session.query(
        Contagem.id,
        Contagem.lote,
        Contagem.validade_mes,
        Contagem.validade_ano,
        Contagem.quantidade,
        Produto.codigo,
        Produto.nome
    ).join(Produto).all()
    
    resultado = []
    for contagem in contagens:
        resultado.append({
            'id': contagem.id,
            'lote': contagem.lote,
            'validade_mes': contagem.validade_mes,
            'validade_ano': contagem.validade_ano,
            'quantidade': contagem.quantidade,
            'produto': {
                'codigo': contagem.codigo,
                'nome': contagem.nome
            }
        })
    
    return jsonify(resultado)

@contagem_bp.route('/contagens', methods=['POST'])
def registrar_contagem():
    """Registra uma nova contagem ou atualiza quantidade se lote já existe"""
    data = request.get_json()
    
    if not data or not all(k in data for k in ['produto_codigo', 'lote', 'validade_mes', 'validade_ano', 'quantidade']):
        return jsonify({'erro': 'Todos os campos são obrigatórios: produto_codigo, lote, validade_mes, validade_ano, quantidade'}), 400
    
    # Buscar produto pelo código
    produto = Produto.query.filter_by(codigo=data['produto_codigo']).first()
    if not produto:
        return jsonify({'erro': 'Produto não encontrado'}), 404
    
    # Verificar se já existe contagem para este produto e lote
    contagem_existente = Contagem.query.filter_by(
        produto_id=produto.id,
        lote=data['lote']
    ).first()
    
    try:
        if contagem_existente:
            # Somar à quantidade existente
            contagem_existente.quantidade += int(data['quantidade'])
            # Atualizar validade se fornecida
            contagem_existente.validade_mes = int(data['validade_mes'])
            contagem_existente.validade_ano = int(data['validade_ano'])
            db.session.commit()
            return jsonify({
                'mensagem': 'Quantidade adicionada ao lote existente',
                'contagem': contagem_existente.to_dict()
            })
        else:
            # Criar nova contagem
            nova_contagem = Contagem(
                produto_id=produto.id,
                lote=data['lote'].strip(),
                validade_mes=int(data['validade_mes']),
                validade_ano=int(data['validade_ano']),
                quantidade=int(data['quantidade'])
            )
            db.session.add(nova_contagem)
            db.session.commit()
            return jsonify({
                'mensagem': 'Nova contagem registrada',
                'contagem': nova_contagem.to_dict()
            }), 201
    except ValueError:
        return jsonify({'erro': 'Valores numéricos inválidos'}), 400
    except Exception as e:
        db.session.rollback()
        return jsonify({'erro': f'Erro interno: {str(e)}'}), 500

@contagem_bp.route('/contagens/produto/<produto_codigo>', methods=['GET'])
def listar_contagens_produto(produto_codigo):
    """Lista todas as contagens de um produto específico"""
    produto = Produto.query.filter_by(codigo=produto_codigo).first()
    if not produto:
        return jsonify({'erro': 'Produto não encontrado'}), 404
    
    contagens = Contagem.query.filter_by(produto_id=produto.id).all()
    return jsonify([contagem.to_dict() for contagem in contagens])

@contagem_bp.route('/contagens/<int:contagem_id>', methods=['PUT'])
def atualizar_contagem(contagem_id):
    """Atualiza uma contagem específica"""
    contagem = Contagem.query.get_or_404(contagem_id)
    data = request.get_json()
    
    if not data:
        return jsonify({'erro': 'Dados não fornecidos'}), 400
    
    try:
        if 'lote' in data:
            contagem.lote = data['lote'].strip()
        if 'validade_mes' in data:
            contagem.validade_mes = int(data['validade_mes'])
        if 'validade_ano' in data:
            contagem.validade_ano = int(data['validade_ano'])
        if 'quantidade' in data:
            contagem.quantidade = int(data['quantidade'])
        
        db.session.commit()
        return jsonify(contagem.to_dict())
    except ValueError:
        return jsonify({'erro': 'Valores numéricos inválidos'}), 400
    except IntegrityError:
        db.session.rollback()
        return jsonify({'erro': 'Lote já existe para este produto'}), 409

@contagem_bp.route('/contagens/<int:contagem_id>', methods=['DELETE'])
def deletar_contagem(contagem_id):
    """Deleta uma contagem específica"""
    contagem = Contagem.query.get_or_404(contagem_id)
    db.session.delete(contagem)
    db.session.commit()
    return jsonify({'mensagem': 'Contagem deletada com sucesso'})

@contagem_bp.route('/relatorio/resumo', methods=['GET'])
def relatorio_resumo():
    """Gera resumo para relatório com totais por produto"""
    resumo = db.session.query(
        Produto.codigo,
        Produto.nome,
        func.sum(Contagem.quantidade).label('total_quantidade')
    ).join(Contagem).group_by(Produto.id, Produto.codigo, Produto.nome).order_by(Produto.codigo).all()
    
    resultado = []
    total_geral = 0
    
    for item in resumo:
        total_produto = item.total_quantidade or 0
        total_geral += total_produto
        resultado.append({
            'codigo': item.codigo,
            'nome': item.nome,
            'total_quantidade': total_produto
        })
    
    return jsonify({
        'produtos': resultado,
        'total_geral': total_geral
    })

