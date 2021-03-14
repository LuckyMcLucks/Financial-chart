from openpyxl import Workbook,load_workbook
import os

#if __name__ == "__main__":
   # os.chdir('C:/Users/Chua Wei Yang/Desktop/Project/Ouroborus/Data/Finance')
 #   wb = load_workbook('FINANCE.xlsx')
    # grab the active worksheet
 #   ws = wb.active

def openxl_helper(sheet,formula): #take in string

    try:

        if 'SUM' in formula:
            return sum(sheet,formula)
        elif '/' in formula:
            return divide(sheet,formula)
        elif '=' in formula:
            
            ref = sheet[formula[1:]]
            if '=' in ref.value:
                return  openxl_helper(sheet,ref.value)
                
            else:
                return ref.value
        
    except: 
        
        return formula 
    return formula
def divide(sheet,formula):
    
    formula = formula[1:]
    values = formula.split('/')
    print(values)
    total = 0
    for i in values:
        print(i)
        if total ==0 :
            if i ==int or i == float:
                total+= i
            else:
                total += openxl_helper(sheet,'='+i)
        else:
            if i ==int or i == float:
                total= total/ i
            else:
                total= total/openxl_helper(sheet,'='+i)

    return round(total,4)
def sum(sheet,formula):
    Range =formula[formula.find('(')+1:formula.find(')')]

    start , end = Range.split(':')
    SUM = sheet[start:end]
    
    total =0 

    for col in SUM:

        for cell in col: 


            if cell.value ==None:
                pass
            elif type(cell.value) ==int or type(cell.value) == float:
                
                total+=cell.value
            else:
                
                total += openxl_helper(sheet,cell.value)


    return round(total,2)

