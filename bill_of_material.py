import pandas as pd
import xlsxwriter


workbook = xlsxwriter.Workbook('Output BOM File.xlsx')
df = pd.read_excel("BOM file for Data processing.xlsx")

items = df['Item Name'].unique()

bold = workbook.add_format({'bold': True})

for item in items:
    
    global_item = item
    current_item = df[df['Item Name']==global_item]
    levels = current_item['Level'].unique()
    
    #getting all the raw materials that are multilevel
    multi ={}
    for l in range(len(levels)):
        for i in range(len(current_item.index)):
            if(l<len(levels)-1 and i<len(current_item.index)-1):
                if(current_item.iloc[i]['Level']==levels[l] and current_item.iloc[i+1]['Level']==levels[l+1]):
                 
                    multi[levels[l+1]]=current_item.iloc[i]['Raw material']
    level_index=0
    
    for level in levels:
        current_item = df[df['Item Name']==global_item]
        current_item = current_item[current_item['Level']==level]
        
        if(level_index>0):
            item=multi[level] 
            
        worksheet = workbook.add_worksheet(item)
        worksheet.write(0,0,'Finished Good List')
        worksheet.write(1,0,'#',bold)
        worksheet.write(1,1,'Item Description',bold)
        worksheet.write(1,2,'Quantity',bold)
        worksheet.write(1,3,'Unit',bold)
        worksheet.write(2,0,1)
        worksheet.write(2,1,item)
        worksheet.write(2,2,1)
        worksheet.write(2,3,'Pc')
        worksheet.write(3,0,'End of FG')
        worksheet.write(4,0,'Raw Material List')
        worksheet.write(5,0,'#',bold)
        worksheet.write(5,1,'Item Description',bold)
        worksheet.write(5,2,'Quantity',bold)
        worksheet.write(5,3,'Unit',bold)
        row=0
        for i in range(len(current_item.index)): 
            
                worksheet.write(row+6,0,row+1)
                worksheet.write(row+6,1,current_item.iloc[i]['Raw material'])
                worksheet.write(row+6,2,current_item.iloc[i]['Quantity'])
                worksheet.write(row+6,3,current_item.iloc[i]['Unit'])
                row=row+1
                
        level_index=level_index+1
        worksheet.write(row+6,0,'End of RM')
        
            
workbook.close()