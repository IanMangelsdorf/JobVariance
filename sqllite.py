from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, Date






Base = declarative_base()

metadata = MetaData()





class Variance(Base):
    __tablename__= 'variance'

    id = Column(Integer, primary_key=True, nullable=False)
    project_no = Column(String())
    bu = Column(String())
    project_name = Column(String())
    complete = Column(String())
    revenue = Column(String())
    cost = Column(String())
    stage = Column(String())
    report_month = Column(Date())


engine = create_engine('sqlite:///variance_data.db')
Base.metadata.create_all(engine)
