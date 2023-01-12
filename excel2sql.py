import xml.etree.ElementTree as xmltree, sqlite3

counter=int()
rowcount=int()
cellcount=int()
emptyrows=int()
emptycells=int()
i=int()
cellx=list()
insertempty=True # a boolean which if True will insert the empty rows found in the excel file into the database

# CLEANS THE EXCEL .XML FILE AND SAVES INTO A NEW FILE

filexml=input('\nEnter .xml file name: ')
filexmlr=open(filexml,'r') #opens filexml for reading only

filexml=filexml.replace('.xml','-clean.xml') # adds -clean to the file name
filexmlw=open(filexml,'w') # creates a new file for writting clean xml

ignoreemptyrow=input('Ignore empty rows? (Yes/No): ')
if ignoreemptyrow.lower()=='yes': insertempty=False

wbcount=5 #this will count 4 lines to ommit writting them (after <Workbook>)
for line in filexmlr:

   if line.startswith('<Workbook'):  #if True, it will write <Workbook> and ommit 4 lines, deletes schemas xmlns=
      filexmlw.write('<Workbook>'+'\n') 
      wbcount=0 
   if wbcount<5: #a count to ommit writting following 4 lines after <Workbook because they contain schemas xmlns=
      wbcount=wbcount+1
      continue #continues with next line in for. the if inquiry starts at the <Workbook line

   if line.startswith(' <DocumentProperties'): #deletes schemas xmlns=
      filexmlw.write(' <DocumentProperties>'+'\n')
      continue
   if line.startswith(' <OfficeDocumentSettings'): #deletes schemas xmlns=
      filexmlw.write(' <OfficeDocumentSettings>'+'\n')
      continue
   if line.startswith(' <ExcelWorkbook'): #deletes schemas xmlns=
      filexmlw.write(' <ExcelWorkbook>'+'\n')
      continue
   if line.startswith('  <WorksheetOptions'): #deletes schemas xmlns=, 1 for every Excel Sheet
      filexmlw.write('  <WorksheetOptions>'+'\n')
      continue
   line=line.replace('ss:','') #deletes ss: schema properties
   line=line.replace('x:','') #deletes x: schema properties
   filexmlw.write(line)

filexmlr.close()
filexmlw.close() # saves and closes the file with the changes. Before this the file is still empty.

#UPTO HERE IT CREATED A CLEAN XML FILE

#CREATE SQL DATABASE AND TABLE

filesql=filexml.replace('.xml','.sqlite') #use the same xml file name for the database
con=sqlite3.connect(filesql)
cur=con.cursor()

col=input('Enter num. of columns: ') #UPTO 20 COLUMNS
colnum=int(col)

#CREATE TABLE DEPENDING ON THE NUM. OF COLUMNS

cur.execute('drop table if exists Table1')

sqlcommand='create table Table1(value0 TEXT'
for i in range(colnum-1):
   sqlcommand=sqlcommand+',value'+str(i+1)+' TEXT'
sqlcommand=sqlcommand+')'

cur.execute(sqlcommand)

#CREATE AN EMPTY LIST cellx[] WITH SIZE COLNUM

i=0 #it's a counter to create an empty list of size colnum
while i<colnum:
   cellx.append('')
   i=i+1

#CREATE SQL COMMAND FOR INSERTING ROWS WITH VARIABLE COLUMNS NUMBER

sqlcommand='insert into Table1(value0'
for i in range(colnum-1):
   sqlcommand=sqlcommand+',value'+str(i+1)
sqlcommand=sqlcommand+') values (?'
for i in range(colnum-1):
   sqlcommand=sqlcommand+',?'
sqlcommand=sqlcommand+')'

#DEFINE FUNCTION INSERTROW() TO INSERT VALUES INTO DB

def insertrow():
   if colnum==1:
      cur.execute(sqlcommand,(cellx[0],)) #empty row, 1 col. you pass a tuple anyway (cellx[0],)
   if colnum==2:
      cur.execute(sqlcommand,(cellx[0],cellx[1])) #empty row, 2 col
   if colnum==3:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2])) #empty row, 3 col, ETC
   if colnum==4:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3]))
   if colnum==5:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4]))
   if colnum==6:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5]))
   if colnum==7:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6]))
   if colnum==8:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7]))
   if colnum==9:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8]))
   if colnum==10:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9]))
   if colnum==11:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10]))
   if colnum==12:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11]))
   if colnum==13:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12]))
   if colnum==14:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13]))
   if colnum==15:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13],cellx[14]))
   if colnum==16:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13],cellx[14],cellx[15]))
   if colnum==17:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13],cellx[14],cellx[15],cellx[16]))
   if colnum==18:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13],cellx[14],cellx[15],cellx[16],cellx[17]))
   if colnum==19:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13],cellx[14],cellx[15],cellx[16],cellx[17],
                 cellx[18]))
   if colnum==20:
      cur.execute(sqlcommand,(cellx[0],cellx[1],cellx[2],cellx[3],cellx[4],cellx[5],cellx[6],cellx[7],cellx[8],
                 cellx[9],cellx[10],cellx[11],cellx[12],cellx[13],cellx[14],cellx[15],cellx[16],cellx[17],
                 cellx[18],cellx[19]))

#ANALYZE XML FILE AND RETRIEVE ALL ROWS

xmldata=xmltree.parse(filexml)
rows=xmldata.findall('Worksheet/Table/Row') # rows is a list of tags Row (all rows of the xml file). row is each individual row.

#EXCEL XML STRUCTURE
#<workbook>
#   <Worksheet>
#      <Table>
#         <Row>
#            <Cell>
#               <Data>

# NOTE: A ROW NOT NECESSARILY HAS DATA, IT CAN HAVE ONLY STYLE FORMAT FOR EXAMPLE. ONLY ROWS WITH DATA WILL BE COUNTED.
# EXCEL USES AN ATTRIBUTE INDEX TO INFORM THAT THE ROW HAS EMPTY ROWS BEFORE (EMPTY OF DATA AND FORMAT STYLE, ETC.) AND 
# AVOIDS MAKING A ROW TAB FOR AN EMPTY ROW.

rowcount=0 #counts rows (tag Row) in the xml file
countempty=0 #counts total empty rows
countrowdata=0 # counts rows with data
print('Inserting data into sqlite database...')

for row in rows: # rows is the group of all rows of the xml file. row is each individual row.

   founddata=False

# DELETE CELL LIST VALUES. THIS AVOIDS INSERTING VALUES FROM PREVIOUS ROWS

   i=0
   while i<colnum:
      cellx[i]='' # overwrite previous values with '', instead of doing append. The list is already created.
      i=i+1

#CHECK EMPTY ROWS AND INSERT THEM INTO DATABASE
   
   rowcount=rowcount+1 #counts rows (tag Row) in the xml file
   rowindex=row.get('Index')
   if rowindex is not None: #that means attribute Index is present, and a number. REVIEW ADD A BOOLEAN TO LET THE USER DECIDE TO INSERT EMPTY ROWS OR NOT
       emptyrows=int(rowindex)-rowcount #will give the number of empty lines to insert
       if insertempty is True: # only if the user choses to insert empty rows do it
          i=0 #counts the num. of empty fields to enter
          while i<emptyrows: #cellx[] is going to be empty ''
             insertrow()
             i=i+1
       rowcount=rowcount+emptyrows # update rowcount
       countempty=countempty+emptyrows

#LOOK FOR TAGS CELL AND INSERT ROW INTO DATABASE

   cellcount=0 #counts the cells/fields in every row
   counter=0 #index with a size upto the num. of columns
   cells=row.findall('Cell') #makes a list of tags Cell in each individual row
   for cell in cells:
      cellcount=cellcount+1
      cellindex=cell.get('Index') # if it has the attribute Index it means there were empty cells before this cell
      if cellindex is not None: #that means attribute Index is present in Cell, and a number
         emptycells=int(cellindex)-cellcount #calculates how many empty cells to insert
         counter=counter+emptycells # adds emptycells to counter which is the cellx[] index, to skip cells
         cellcount=cellcount+emptycells #when detecting empty cells, add them to cellcount keep updated
      if cell.find('Data') is not None:
         cellx[counter]=cell.find('Data').text #stores the text value inside the cellx at index counter
         founddata=True # only if a cell has tag data will it count that row as having data
      else:
         pass #if there is no tag Data inside Cell ommit the cell, it's empty. E.g. Some cells might only have style format.
      counter=counter+1
   if insertempty is True:
      insertrow()
      if founddata is True:
         countrowdata=countrowdata+1
   else:
      if founddata is True:
         insertrow()
         countrowdata=countrowdata+1

lenrows=len(rows) # total tags Row found by the xml parser
print('Total rows:',(lenrows+countempty)) # empty rows without style format won't appear in the xml file as rows. They have to be counted manually.
print('Rows with data:',countrowdata)
print('Empty rows:',countempty)
print('Empty rows with format:',(lenrows-countrowdata))

con.commit()