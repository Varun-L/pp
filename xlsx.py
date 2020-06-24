from xlsxwriter import Workbook as wb

w1 = wb('sample.xlsx')

ws=w1.add_worksheet()

bold=w1.add_format({'bold':True})
money=w1.add_format({'num_format':'$#,##'})

ws.write('A1','Item',bold)
ws.write('B1','Cost',bold)

ex = (
    ['Rent', 1000],
    ['Gas', 100],
    ['Food', 300],
    ['Gym', 50],
)

r1=1
c1=0
for i,c in ex:
    ws.write(r1,c1,i)
    ws.write(r1,c1+1,c)
    r1+=1
ws.write(r1,0,'Total',bold)
ws.write(r1,1,'=SUM(B2:B5)',money)

w1.close()

