import sqlalchemy as db

engine = db.create_engine('sqlite:///VPN.sqlite')
connection = engine.connect()
metadata = db.MetaData()


Customer = db.Table('Customer', metadata,
	db.Column('name', db.String(255), primary_key=True, nullable=False),
	db.Column('pm', db.String(255), nullable=False),
	db.Column('phone', db.String(255), default=),
	db.Column('email', db.String(255), default=True)
)

metadata.create_all(engine)