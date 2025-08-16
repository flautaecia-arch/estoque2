from flask_sqlalchemy import SQLAlchemy
from src.models.user import db

class Contagem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    produto_id = db.Column(db.Integer, db.ForeignKey('produto.id'), nullable=False)
    lote = db.Column(db.String(100), nullable=False)
    validade_mes = db.Column(db.Integer, nullable=False)
    validade_ano = db.Column(db.Integer, nullable=False)
    quantidade = db.Column(db.Integer, nullable=False, default=0)
    
    # Índice único para evitar duplicação de lotes por produto
    __table_args__ = (db.UniqueConstraint('produto_id', 'lote', name='unique_produto_lote'),)

    def __repr__(self):
        return f'<Contagem Produto:{self.produto_id} Lote:{self.lote} Qtd:{self.quantidade}>'

    def to_dict(self):
        return {
            'id': self.id,
            'produto_id': self.produto_id,
            'lote': self.lote,
            'validade_mes': self.validade_mes,
            'validade_ano': self.validade_ano,
            'quantidade': self.quantidade
        }

