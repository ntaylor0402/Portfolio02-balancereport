import openpyxl

def getColumns(ws):
    """traverses tables column title row and returns a dictionary with column titles and column index key-pair values"""

    columns = {}

    for i in range(1, ws.max_column + 1):
        thisCell = ws.cell(row=1, column=i)
        columns.update({thisCell.value: thisCell.column})
    return columns





