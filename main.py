from openpyxl import load_workbook

class ARRLStates:
    def __init__(self, row):
        self.State = row[0].value.rstrip().lstrip()
        self.Eighty = row[1].value.rstrip().lstrip()
        self.Fourty = row[2].value.rstrip().lstrip()
        self.Twenty = row[3].value.rstrip().lstrip()
        self.Fifteen = row[4].value.rstrip().lstrip()
        self.Ten = row[5].value.rstrip().lstrip()

wb = load_workbook(filename = 'ARRL.xlsx')
sheet_ranges = wb['Sheet1']
States = []

for row in sheet_ranges.rows:
    x = ARRLStates(row)
    States.append(x)
bands = ['Eighty', 'Fourty', 'Twenty', 'Fifteen', 'Ten']

def sorted(x):
    return [s.State for s in States if not s.__getattribute__(x)]

for b in bands:
    a = sorted(b)
    print(b, len(a), '\n', a)
