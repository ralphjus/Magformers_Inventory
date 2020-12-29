import openpyxl
from openpyxl import load_workbook
from square.client import Client
import datetime 
from datetime import timezone
import uuid
import json

def pull_from_square():

    wb = openpyxl.load_workbook('Inventory.xlsx')
    ws = wb.active
    wb.sheetnames

    product_line = input("Please input product line to pull (tab of excel):\n")
    sheet = wb[product_line]

    item_ids = []
    for item in range(4,sheet.max_row +1):
        item_ids.append(sheet['F'+str(item)].value)

    client = Client(
        access_token='ACCESSTOKEN',
        environment='production',
    )
    objects = []
    quantities = []
    for item in item_ids:
        if item != None:
            result = client.inventory.retrieve_inventory_count(
            catalog_object_id = item,
            location_ids = "CMRZVV9YNFAJP"
        )

            if result.is_success():
                json_object = json.dumps(result.body, indent = 4)
                try:
                    data = json.loads(json_object)
                    counts = data['counts']
                    object_id = counts[0]['catalog_object_id']
                    quantity = counts[0]['quantity']
                    objects.append(object_id)
                    quantities.append(quantity)
                except:
                    print("oops!!!")
                    objects.append(0)
                    quantities.append(0)
                
            elif result.is_error():
              print(result.errors)
          
        else:
            objects.append(0)
            quantities.append(0)
        
    for item in range(len(objects)):
        if objects[item] in item_ids:
            index = int(item_ids.index(objects[item]))
            row = index + 4
            sheet['E'+str(row)] = int(quantities[item])
            print("{} is now {}".format(sheet['B'+str(row)].value,quantities[item]))
    wb.save('Inventory.xlsx')
    
if __name__ == '__main__':
    print("You should run me in main, but ok!")
    pull_from_square()
