import openpyxl as op

work = op.Workbook()
sht = work.active
sht['a1'] = "Oreder ID"
sht['b1'] = "Customer Name"
sht['c1'] = "Product"
sht['d1'] = "Quantity"
sht['e1'] = "Price"
sht['f1'] = "Total"


sht['a2'] = "1"
sht['b2'] = "Juan Dela Cruz"
sht['c2'] = "Burger"
sht['d2'] = "2"
sht['e2'] = "75"
sht['f2'] = "150"

sht['a3'] = "2"
sht['b3'] = "Maria Santos"
sht['c3'] = "Fries"
sht['d3'] = "3"
sht['e3'] = "50"
sht['f3'] = "150"
sht['a3'] = "2"

sht['a4'] = "3"
sht['b4'] = "Carlos Reyes"
sht['c4'] = "Pizza"
sht['d4'] = "1"
sht['e4'] = "350"
sht['f4'] = "350"

sht['a5'] = "4"
sht['b5'] = "Angela Lopez"
sht['c5'] = "Milktea"
sht['d5'] = "4"
sht['e5'] = "120"
sht['f5'] = "480"

sht['a6'] = "5"
sht['b6'] = "Kevin Ramos"
sht['c6'] = "Spaghetti"
sht['d6'] = "2"
sht['e6'] = "95"
sht['f6'] = "190"
work.save("ordersDB.xlsx")