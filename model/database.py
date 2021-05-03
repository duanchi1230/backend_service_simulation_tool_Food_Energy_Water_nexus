"""
This module is obsolete and not used in FEWSim
"""
import sqlalchemy as db
from sqlalchemy import Column, Integer, String, JSON, Float, ARRAY
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import json
engine = db.create_engine('sqlite:///scenarios.db', echo=True)
Base = declarative_base()

print(input)
class Scenarios(Base):
	__tablename__ = 'scenarios'
	id = Column(Integer, primary_key=True)
	name = Column(String)
	WEAP_Input = Column(JSON)
	LEAP_Input = Column(JSON)
	MABIA_Input = Column(JSON)

	WEAP_Output = Column(JSON)
	LEAP_Output = Column(JSON)

	def __repr__(self):
		return "<User(scenario_name='%s')>" % (
			self.name)

Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()
# ed_user = Scenarios(name='5% population growth', WEAP_Input={"population growth":5, "municipal efficiency":10})
# session.add(ed_user)
# session.commit()

# session.query(Scenarios).delete()
# session.commit()

# print(session.query(Scenarios).filter_by(name='5% population growth').all()[0].WEAP_Input)
