from openpyxl import load_workbook

wb = load_workbook(filename = 'ARRL.xlsx')
class ARRLStates:
    def __init__(self, row):
        self.State = row[0].value.rstrip().lstrip()
        self.Eighty = row[1].value
        self.Fourty = row[2].value
        self.Twenty = row[3].value
        self.Fifteen = row[4].value
        self.Ten = row[5].value

sheet_ranges = wb['Sheet1']
States = []
for row in sheet_ranges.rows:
    x = ARRLStates(row)
    States.append(x)
    
state_bands = ['Eighty', 'Fourty', 'Twenty', 'Fifteen', 'Ten']
def sorted(x):
    return [s.State for s in States if not s.__getattribute__(x)]

for b in state_bands:
    a = sorted(b)
    print(b, len(a), '\n', a)

class DXCountries:
    def __init__(self, row):
        self.Country = row[0].value.rstrip().lstrip()
        self.OneSixty = row[1].value
        self.Eighty = row[2].value
        self.Fourty = row[3].value
        self.Thirty = row[4].value
        self.Twenty = row[5].value
        self.Seventeen = row[6].value
        self.Fifteen = row[7].value
        self.Twelve = row[8].value 
        self.Ten = row[9].value
        self.Six = row[10].value

dxcc_bands = ['OneSixty', 'Eighty', 'Fourty', 'Thirty', 'Twenty', 'Seventeen', 'Fifteen', 'Twelve', 'Ten', 'Six']
def sorted_countries(x):
    return [s.Country for s in dxcc if s.__getattribute__(x)]

sheet_ranges = wb['Sheet2']
dxcc = []

for row in sheet_ranges.rows:
    if row[0].value is not None:
        x = DXCountries(row)
        dxcc.append(x)

for d in dxcc_bands:
    a = sorted_countries(d)
    print(d, len(a), '\n', a)
