from collections import defaultdict
# In your python terminal cd into the directory having the excel file. 
import xlrd
import time
from prettytable import PrettyTable

# INPUT DATA
book = xlrd.open_workbook(input('Enter the name of the cost and volume workbook : ')+".xlsx")
sheet = book.sheet_by_name(input('Enter the name of the cost and volume sheet: '))
constraintbook = xlrd.open_workbook(input('Enter the name of the constraint workbook : ')+".xlsx")
constraint = constraintbook.sheet_by_name(input('Enter the name of the constraint sheet: '))
colnum = sheet.ncols-1
rownum = sheet.nrows-1
costs  = {}
eachcost = {}
for c in range(1,colnum):
    for r in range (1,rownum):
        if constraint.cell(r,c).value == 1:
            eachcost[sheet.cell(r,0).value]=sheet.cell(r,c).value
        else: eachcost[sheet.cell(r,0).value]=sheet.cell(r,c).value*10000    
    costs[sheet.cell(0,c).value]=eachcost.copy()
demand = {}
for r in range (1,rownum):
    demand[sheet.cell(r,0).value]=sheet.cell(r,colnum).value   
supply = {}
for c in range(1,colnum):
    supply[sheet.cell(0,c).value]=sheet.cell(rownum,c).value 
cols = sorted(demand.keys())
res = dict((k, defaultdict(int)) for k in costs)
g = {}
for x in supply:
    g[x] = sorted(costs[x].keys(), key=lambda g: costs[x][g])
for x in demand:
    g[x] = sorted(costs.keys(), key=lambda g: costs[g][x])

# MAIN ALGORITHM 
while g:
    d = {}
    for x in demand:
        d[x] = (costs[g[x][1]][x] - costs[g[x][0]][x]) if len(g[x]) > 1 else costs[g[x][0]][x]
    s = {}
    for x in supply:
        s[x] = (costs[x][g[x][1]] - costs[x][g[x][0]]) if len(g[x]) > 1 else costs[x][g[x][0]]
    f = max(d, key=lambda n: d[n])
    t = max(s, key=lambda n: s[n])
    t, f = (f, g[f][0]) if d[f] > s[t] else (g[t][0], t)
    v = min(supply[f], demand[t])
    res[f][t] += v
    demand[t] -= v
    if demand[t] == 0:
        for k in supply:
            if supply[k] != 0:
                g[k].remove(t)
        del g[t]
        del demand[t]
    supply[f] -= v
    if supply[f] == 0:
        for k in demand:
            if demand[k] != 0:
                g[k].remove(f)
        del g[f]
        del supply[f]
    if len(demand)==0 and len(supply)!=0:
        supply.clear()
        g.clear()
            
# CREATE OUTPUT TABLE        
x= PrettyTable()
PL=['Route']
for pl in sorted(costs):
    PL.append(pl) #3PL
x.field_names=PL.copy()
cost = 0
for route in cols: #cols = routes
    row=[route] #3PL
    for n in sorted(costs):
        y = res[n][route]
        row.append(y) #COSTS
        cost += y * costs[n][route]
    x.add_row(row.copy())
print (x)
percent = 0
for f in res:
    for t in res[f]:
        percent += res[f][t]
    print (f+' '+str(round((percent/sheet.cell(rownum,colnum).value)*100,3))+'%')
    percent=0
print ("\n\nTotal Cost = ", cost)
print ('time taken:'+time.clock())
