from flask import Blueprint, request, jsonify
from src.models.user import db
from src.models.produto import Produto
from sqlalchemy.exc import IntegrityError

produto_bp = Blueprint('produto', __name__)

@produto_bp.route('/produtos', methods=['GET'])
def listar_produtos():
    """Lista todos os produtos cadastrados"""
    produtos = Produto.query.all()
    return jsonify([produto.to_dict() for produto in produtos])

@produto_bp.route('/produtos', methods=['POST'])
def criar_produto():
    """Cria um novo produto"""
    data = request.get_json()
    
    if not data or 'codigo' not in data or 'nome' not in data:
        return jsonify({'erro': 'Código e nome são obrigatórios'}), 400
    
    try:
        produto = Produto(
            codigo=data['codigo'].strip(),
            nome=data['nome'].strip()
        )
        db.session.add(produto)
        db.session.commit()
        return jsonify(produto.to_dict()), 201
    except IntegrityError:
        db.session.rollback()
        return jsonify({'erro': 'Produto com este código já existe'}), 409

@produto_bp.route('/produtos/<codigo>', methods=['GET'])
def buscar_produto_por_codigo(codigo):
    """Busca um produto pelo código"""
    produto = Produto.query.filter_by(codigo=codigo).first()
    if not produto:
        return jsonify({'erro': 'Produto não encontrado'}), 404
    return jsonify(produto.to_dict())

@produto_bp.route('/produtos/<int:produto_id>', methods=['PUT'])
def atualizar_produto(produto_id):
    """Atualiza um produto existente"""
    produto = Produto.query.get_or_404(produto_id)
    data = request.get_json()
    
    if not data:
        return jsonify({'erro': 'Dados não fornecidos'}), 400
    
    try:
        if 'codigo' in data:
            produto.codigo = data['codigo'].strip()
        if 'nome' in data:
            produto.nome = data['nome'].strip()
        
        db.session.commit()
        return jsonify(produto.to_dict())
    except IntegrityError:
        db.session.rollback()
        return jsonify({'erro': 'Produto com este código já existe'}), 409

@produto_bp.route('/produtos/<int:produto_id>', methods=['DELETE'])
def deletar_produto(produto_id):
    """Deleta um produto e suas contagens"""
    produto = Produto.query.get_or_404(produto_id)
    db.session.delete(produto)
    db.session.commit()
    return jsonify({'mensagem': 'Produto deletado com sucesso'})

