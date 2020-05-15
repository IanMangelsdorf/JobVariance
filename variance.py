import os
import calendar
import xlrd

from sqllite import Variance, Base
from datetime import datetime
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker


engine = create_engine('sqlite:///variance_data.db')

Base.metadata.bind = engine

DBSession = sessionmaker(bind=engine)
session = DBSession()


rootDir='C:\\Users\\ian.mangelsdorf\\OneDrive - Downer\\Power BI Files\\25.1 Job Variance Reports'

def getdate(str):
    a = str.split(' ')[0]
    b= a.split('_')
    yr = '20' + b[0]
    m = int(b[1])
    mnth = (calendar.month_name[m])
    dt = datetime.strptime(' '.join([mnth,yr]), '%B %Y')
    return dt

def getproject(str):
    if '-' in str:
        a = str.split('-')
        if len(a)==2 and len(a[0])>4:
            return [a[0], a[1]]
        else:
            return ['skip', 'skip']
    else:
        return ['skip','skip']

def gettype(str):
    a = str.split('\n')
    typ = a[-1]
    if typ in ['Revenue', 'Cost', 'Margin']:
        return [a[-1], ' '.join(a[:-1]).strip()]
    else:
        return ['skip', 'skip']

def parse_data (ws,  report_date):
    c=2
    for r in range(ws.nrows):
        pn, project = getproject(ws.cell(r,0).value)
        if not pn == 'skip':
            complete = ws.cell(r,1).value
            bu = ws.cell(r,27).value

            for c in range(ws.ncols):
                if c == 22:
                    state, typ = 'TD','Billings'
                elif c==23:
                    state, typ = 'Balance','WIP'
                elif c==24:
                    state, typ = 'Movement','WIP'
                else:
                    typ, state = gettype(ws.cell(0, c).value)
                if state !='skip' or typ !='skip':
                    new_var = Variance(
                        project_no=pn,
                        bu = bu,
                        project_name = project,
                        complete = complete,
                        revenue = typ,
                        cost = ws.cell(r,c).value,
                        stage = state,
                        report_month = report_date,
                    )

                    session.add(new_var)
                    session.commit()




def main():
    for dirName, subdirList, fileList in os.walk(rootDir):
        for fname in fileList:
            if '.xls' in fname or '.xlsx' in fname:
                report_date = getdate(fname)
                rs = xlrd.open_workbook(os.path.join(dirName, fname))
                parse_data(rs.sheet_by_index(0),  report_date)



if __name__ == '__main__':
    main()

