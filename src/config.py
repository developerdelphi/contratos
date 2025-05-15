import os
basedir = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = os.environ.get('FLASK_SECRET_KEY', 'O130m@rB@Din@1v@l1d@')
    UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER','uploads')
    CREDENTIALS_FILE = os.path.join(basedir, 'utils', 'credentials-google.json')
    
    COLUNA_NOME_PADRAO = 'NOME'
    COLUNA_CPF_PADRAO = 'CPF'

    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + os.path.join(basedir, 'flask_session.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SESSION_TYPE = 'sqlalchemy'
    SESSION_PERMANENT = False
    SESSION_USE_SIGNER = True
    SESSION_SQLALCHEMY_TABLE = 'sessions'
    # ... outras configs ...
