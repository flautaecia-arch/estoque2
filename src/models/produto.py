from flask_sqlalchemy import SQLAlchemy
from src.models.user import db

class Produto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(50), unique=True, nullable=False)
    nome = db.Column(db.String(200), nullable=False)
    
    # Relacionamento com contagens
    contagens = db.relationship('Contagem', backref='produto', lazy=True, cascade='all, delete-orphan')

    def __repr__(self):
        return f'<Produto {self.codigo} - {self.nome}>'

    def to_dict(self):
        return {
            'id': self.id,
            'codigo': self.codigo,
            'nome': self.nome
        }

