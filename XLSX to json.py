import xlrdbook = xlrd.open_workbook("EPON2.xlsx")
sheetcount = book.nsheets
print("done")

f = open("demofile2.json", "w")
sheetcount = 1
for I in range(sheetcount):
    currsheet = book.sheet_by_index(I)
    f.write("{")
    f.write("\n")
    f.write(" \"index\""+":" +" "+  str(I+1) + ",")
    f.write("\n")
    f.write(" \"title\"" + ":" + " " + " \"" + str(currsheet.cell(0,1).value) + "\"" + ",")
    f.write("\n")
    f.write(" \"content\"" + ": " + " \""+str(currsheet.cell(1,1).value) + "\" ")
    f.write("\n")
    labels = list
    f.write(" \"labels\"" + ": " + "["+ "\"" +str(currsheet.cell(1,2).value) + "\""+ "]")
    f.write("\n")
    f.write(" \"subheaders\"" + ": " + "[")
    f.write("\n")
    rows = currsheet.nrows
    cols = currsheet.ncols
    #print(rows)
    #print(cols)
    counter = 0
    for i in range(1,rows-1,2):
        counter = counter+1
        f.write("   {")
        f.write("\n")
        f.write("    \"index\""+":" +" "+  str(counter) + ",")
        f.write("\n")
        f.write("    \"title\"" + ":" + " " + " \"" + str(currsheet.cell(i+1,1).value) + "\"" + ",")
        f.write("\n")
        #print("    \"content\":" + " \"" + str(currsheet.cell(i+2,1).value) + "\"" )
        f.write("    \"content\":" + " \"" + str(currsheet.cell(i+2,1).value) + "\"" )
        f.write("\n")   
        if currsheet.cell(i+2,2).value == "": 
            f.write("    \"labels\"" + ": " + "[]")
        else:
            f.write("    \"labels\"" + ": " + "[" + "\"" + str(currsheet.cell(i+2,2).value) + "\""+ "]")
        f.write("\n")
        f.write("   },")
        f.write("\n")
    counter = 0
    
    f.write("  ]")
    f.write("\n")
    f.write("},")
    f.write("\n")
f.close()

