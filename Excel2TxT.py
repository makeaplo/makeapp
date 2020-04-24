import xlrd
tb=xlrd.open_workbook(r'm:/1.xlsx')
sht=tb.sheets()[0]
for x in range(0,sht.nrows):
    with open(r'm:/%s.txt' % sht.cell_value(x,0),'a') as f:
        f.write(sht.cell_value(x,1)+''+sht.cell_value(x,5)+''+sht.cell_value(x,6)+'\n')
