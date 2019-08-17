import sqlalchemy as db
from sqlalchemy import Column, Integer, String
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from sqlalchemy.orm import relationship
from sqlalchemy import ForeignKey
from sqlalchemy.orm import aliased

print(db.__version__)
engine = db.create_engine('sqlite:///test.db', echo=True)
Base = declarative_base()


class User(Base):
	__tablename__ = 'users'
	id = Column(Integer, primary_key=True)
	name = Column(String)
	fullname = Column(String)
	nickname = Column(String)

	def __repr__(self):
		return "<User(name='%s', fullname='%s', nickname='%s')>" % (
			self.name, self.fullname, self.nickname)

Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()
ed_user = User(name='ed', fullname='Ed Jones', nickname='edsnickname')
session.add(ed_user)

# our_user = session.query(User).filter_by(name='ed').first()
#
# session.add_all([User(name='wendy', fullname='Wendy Williams', nickname='windy'),
#                  User(name='mary', fullname='Mary Contrary', nickname='mary'),
#                  User(name='fred', fullname='Fred Flintstone', nickname='freddy')])
# ed_user.name = 'Edwardo'
# fake_user = User(name='fakeuser', fullname='Invalid', nickname='12345')
# session.add(fake_user)

# for instance in session.query(User).order_by(User.id):
# 	print(instance.id, instance.name, instance.fullname)

# for row in session.query(User.name.label('name_label')).all():
# 	print(row, row.name_label)

# user_alias = aliased(User, name='user_alias')
# for row in session.query(user_alias, user_alias.name).all():
# 	print(row.user_alias)




class Address(Base):
	__tablename__ = 'addressess'

	id = Column(Integer, primary_key=True)
	email_address = Column(String, nullable=False)
	user_id = Column(Integer, ForeignKey('users.id'))

	user = relationship("User", back_populates="a")

	def __repr__(self):
		return "<Address(email_address='%s')>" % self.email_address

User.a = relationship(
	"Address", order_by=Address.id, back_populates="user")

# Base.metadata.create_all(engine)
#
for u in session.query(User).order_by(User.id)[0:3]:
	print(u.a)
print(Address.user)
