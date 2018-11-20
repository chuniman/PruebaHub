import xlwt
import xlrd
import datetime
#ArchivoLeer
fileLocation="C:\\Users\\Usuario\\Documents\\DatosRecursos.xlsx"
workbook=xlrd.open_workbook(fileLocation)
sheet=workbook.sheet_by_index(0)
#Escribir
wb = xlwt.Workbook()
ws=wb.add_sheet("My Sheet",True)

nombreActual=""

for i in range(1,1080):        
        #nombre DataSet
        ws.write(i,0,sheet.cell_value(i,0))
        #nombre Recurso
        ws.write(i,1,sheet.cell_value(i,1))
        #Organismo
        if sheet.cell_value(i,2)!="":
                ws.write(i,2,"Si")
        else:
                ws.write(i,2,"No")
        #fecha
        ws.write(i,3,str(datetime.datetime.now()))
        #Tiene Descripcion
        if sheet.cell_value(i,3)!="":
                ws.write(i,5,"Si")
        else:
                ws.write(i,5,"No")
        #Tiene Descripcion el recurso        
        if sheet.cell_value(i,13)!="":
                ws.write(i,21,"Si")
        else:
                ws.write(i,21,"No")             
        #Cantidad de Recursos de metadatos
        ws.write(i,8,5)
        #Los recursos de metadatos se corresponden con los de datos?
        ws.write(i,9,"Si")
        #La calidad en la descripcion de los metadatos permite comprender los datos en el archivo de datos
        ws.write(i,10,"Si")
        #Tiene autor
        if sheet.cell_value(i,5)!="":
                ws.write(i,11,"Si")
        else:
                ws.write(i,11,"No")
        #Tiene mantenedor
        if sheet.cell_value(i,6)!="":
                ws.write(i,12,"Si")
        else:
                ws.write(i,12,"No")
        #Esta Activo
        if sheet.cell_value(i,18)=="active":
                ws.write(i,13,"Si")
        else:
                ws.write(i,13,"No")
        #Esta actualizado(origen dice outdated)
        if sheet.cell_value(i,19)=="True":
                ws.write(i,14,"No")
        else:
                ws.write(i,14,"Si")
        #Es org del estado, creo
        ws.write(i,15,"si")
        #SI es del estado: tiene licencia DAG-UY
        if sheet.cell_value(i,11)=="Licencia de DAG de Uruguay":
                ws.write(i,16,"Si la tiene")
        else:
                ws.write(i,16,"No la tiene")
        #Si es privado: tiene licencia?(MUST) ,no hay privados
        #Tiene categorias
        if sheet.cell_value(i,10)!="":
                ws.write(i,17,"Si")
        else:
                ws.write(i,17,"No")
        #Tiene visualizaciones
        #Tiene aplicaciones
        if sheet.cell_value(i,22)!="":
                ws.write(i,19,"Si")
        else:
                ws.write(i,19,"No")
        #Si tiene apps: Estan activas
wb.save("myworkbook.xls")           

"""
DATA = (("The Essential Calvin and Hobbes", 1988,),
        ("The Authoritative Calvin and Hobbes", 1990,),
        ("The Indispensable Calvin and Hobbes", 1992,),
        ("Attack of the Deranged Mutant Killer Monster Snow Goons", 1992,),
        ("The Days Are Just Packed", 1993,),
        ("Homicidal Psycho Jungle Cat", 1994,),
        ("There's Treasure Everywhere", 1996,),
        ("It's a Magical World", 1996,),)

wb = xlwt.Workbook()
ws = wb.add_sheet("My Sheet")
for i, row in enumerate(DATA):
    for j, col in enumerate(row):
        ws.write(i, j, col)
ws.col(0).width = 256 * max([len(row[0]) for row in DATA])
wb.save("myworkbook.xls")
"""