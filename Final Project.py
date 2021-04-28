import xlrd
loc = ("/bondData.xls") #reading bondData.xls file
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


#printing all bond price of all the bonds
for i in range(sheet.nrows):
    print(BondPrinceCalculator.bondPrice(sheet.cell_value(i,0),sheet.cell_value(i,0),sheet.cell_value(i,0),sheet.cell_value(i,0),sheet.cell_value(i,0)))


#bond calculator object
class BondPriceCalculator:  
    def __init__(self):
        self.data = None

    def bondPrice(Fv,CR,CF,t,ytm):
        bondprice = ((Fv*CR/CF*(1-(1+ytm/CF)**(-CF*t)))/(ytm/CF)) + Fv*(1+(ytm/CF))**(-CF*t)
        return bondprice


    
    def bondTable():
        bondMatrix = [[0 for j in range(6)] for i in range(21)]
        for i in range(21):
            ytm = 0.07 + i*(0.001)
            bondMatrix[i][0] = 100 
            bondMatrix[i][1] = 0.08
            bondMatrix[i][2] = 1
            bondMatrix[i][3] = 20
            bondMatrix[i][4] = ytm
            bondMatrix[i][5] = bondPrice(100,0.08,1,20,ytm)
        return bondMatrix


#testing
print(BondPriceCalculator.bondPrice(100,0.08,1,20,0.07))



#printing bondTable
print(BondPriceCalculator.bondTable())







#ploting of data
import matplotlib.pyplot as plt

fig, ax =plt.subplots(1,1)
data=bondMatrix
column_labels=["Face Value", "Coupon Rate", "Frequency", "Time to Maturity", "YTM", "Bond Price"]
ax.axis('tight')
ax.axis('off')
ax.table(cellText=data,colLabels=column_labels,loc="center")

plt.show()
