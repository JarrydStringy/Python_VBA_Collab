import xlwings as xw
from xlwings import Book

filePath = "Comparison_Tool.xlsm"

def Sort():
    Book(filePath).set_mock_caller()

    wb = xw.Book.caller()
    ws = wb.sheets[0]

    lRow = ws.range(1,1).end('down').row

    #Sort LHS
    sort1 = ws.range("B2:B{row}".format(row=lRow))      #Will firstly sort by col B
    sort2 = ws.range("C2:C{row}".format(row=lRow))      #Then C
    sort3 = ws.range("D2:D{row}".format(row=lRow))      #Then D
    dataRange = ws.range("B1:D{row}".format(row=lRow))  #The data range that will be sorted
    ws.range(dataRange).api.Sort(                       
        Key1=sort1.api, Order1=1,                       #Run the sort for 1 in ascending order
        Key2=sort2.api, Order2=1,
        Key3=sort3.api, Order3=1,
        Header=1, Orientation=1)                        #Header will ignore the first line

    #Sort RHS
    sort1 = ws.range("H2:H{row}".format(row=lRow))      #Will firstly sort by col B
    sort2 = ws.range("I2:I{row}".format(row=lRow))      #Then C
    sort3 = ws.range("J2:J{row}".format(row=lRow))      #Then D
    dataRange = ws.range("H1:J{row}".format(row=lRow))  #The data range that will be sorted
    ws.range(dataRange).api.Sort(                       
        Key1=sort1.api, Order1=1,                       #Run the sort for 1 in ascending order
        Key2=sort2.api, Order2=1,
        Key3=sort3.api, Order3=1,
        Header=1, Orientation=1)                        #Header will ignore the first line

    # for row in range(2, lRow):


def Clear():
    Book(filePath).set_mock_caller()
    print("Clear")


if __name__ == "__main__":
    print("Here")
    # To be able to easily invoke such code from Python for debugging
    Book(filePath).set_mock_caller()
