dictionary = []
#Set fields to dictionary
def setDictionaryFields(ObjectXls):
    global dictionary
    book = ObjectXls
    sheetToWork = book.sheet_by_index(2)
    cells = sheetToWork.row_slice(rowx=0,
                                  start_colx=0,)

    for cell in reversed(cells):
        dictionary.insert(0, cell.value.lower())

    pass

def getDictionaryFields():
    global dictionary
    return dictionary
    pass
