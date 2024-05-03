import pandas as pd
import math

x1= pd.ExcelFile('DELIVERY.xlsx')
df=x1.parse('28')

for index, row in df.iterrows():
    po_number=row['PO NO']
    if not math.isnan(po_number):
#getting all details from that row
        art=row['ART']
        clr=row['CLR']
        grid=str(row['GRID'])
        if 'X' in grid:
            grid_split=grid.split('X')
            min_grid=int(grid_split[0])
            max_grid=int(grid_split[1])
        else:
            print("GRID TYPO ERROR!!!")


        # checking for all fields and then openting 'stiching' file
        if art and clr and min_grid and max_grid:

            stiching=pd.ExcelFile('STICHING.xlsx')
            sheet=stiching.parse('Sheet1')

            for st_index, st_row in sheet.iterrows():
                if st_row['Pur. Doc.']==po_number:
                    short_text=str(st_row['Short Text'])
                    material=str(st_row['Material'])
                    print("PO NUMBER MATCHING ",st_index)
                    print("Short Text:",short_text)
                    print("ART:",art,"--CLR:",clr)
                    print("*************************************\n--------------------------------")
                    if art in short_text and clr in material:
                        print("ART and CLR is matching ")
                        sheet['Remark'] = sheet['Remark'].astype(float)
                        sheet.loc[st_index, 'Remark'] = 1245

                        sheet.to_excel('STICHING.xlsx', index=False)



