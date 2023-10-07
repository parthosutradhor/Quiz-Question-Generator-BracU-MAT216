import xlrd
from sympy import Matrix, pprint
workbook = xlrd.open_workbook('ATTENDANCE_SHEET.xls')
sheet = workbook.sheet_by_name('DynamicReport')  # Replace 'Sheet1' with the actual sheet name

def add_text_in_file(ID, Name):
    try:
        with open('origin.txt', 'r') as file:
            # Read the content of the file
            file_content = file.read()
        
        
        
        # Replace the old text with the new text
        d1 = int((ID%1000)/100)
        d2 = int((ID%100)/10)
        d3 = (ID%10)
        

        # Given expressions
        expressions = [
            (7*d1 + 3*d2 + 2*d3) % 10 + d2,
            (9*d1 - 5*d2 + 6*d3) % 10 + d1,
            -((4*d1 + 2*d2 - 8*d3) % 10 + d3+1),
            (1*d1 - 9*d2 + 7*d3) % 10 + d1,
            (8*d1 + 6*d2 + 3*d3) % 10 + d3,
            0,
            
            (6*d1 - 7*d2 + 4*d3) % 10 + d2,
            (3*d1 + 5*d2 - 9*d3) % 10 + d1,
            -((5*d1 - 8*d2 + 1*d3) % 10 + d3+1),
            (2*d1 + 7*d2 + 6*d3) % 10 + d1,
            (9*d1 - 4*d2 + 8*d3) % 10 + d2,
            0,
            
            (1*d1 + 9*d2 - 2*d3) % 10 + d1,
            -((4*d1 + 6*d2 + 7*d3) % 10 + d2+1),
            ((8*d1 - 3*d2 + 5*d3) % 10 + d3),
            (2*d1 + 8*d2 + 4*d3) % 10 + d2,
            (7*d1 - 5*d2 + 1*d3) % 10 + d1,
            0,
            
            (3*d1 + 6*d2 + 9*d3) % 10 + d3,
            (5*d1 - 8*d2 + 4*d3) % 10 + d2,
            ((1*d1 + 7*d2 + 3*d3) % 10 + d1),
            (6*d1 - 2*d2 + 5*d3) % 10 + d2,
            -((4*d1 + 9*d2 - 7*d3) % 10 + d1+1),
            0
        ]
        
       

        # Reshape the expressions into a 4x6 matrix
        M_values = [expressions[i:i+6] for i in range(0, len(expressions), 6)]
        M_b = Matrix(M_values)
        
        for k in range(4): 
            M_b[5+6*k] = (d1+1)*M_b[0+6*k] + d2*M_b[1+6*k] - d3*d3*M_b[2+6*k] + (d1-d3)*M_b[3+6*k] + (d1+ d1*d2)*M_b[4+6*k]
        
        if(d3==0 or d3==1 or d3==9):
            for k in range(6):
                M_b[6+k] = d1*M_b[0+k] + d2*M_b[12+k] - d3*M_b[18+k]
        if(d3==0):
            for k in range(6):
                M_b[18+k] = (d3+1)*M_b[0+k] + d2*M_b[12+k] - d1*M_b[6+k]
        
        # Print the matrix
        print("The original System:\n")
        pprint(M_b)   
        print("\n")
        M_rref = M_b.rref()
        print("The Reduced System:\n")
        pprint(M_rref[0])  
        print("\n\n")
        
        

    except FileNotFoundError:
        print(f'The file "{file_path}" was not found.')
    except Exception as e:
        print(f'An error occurred: {str(e)}')



print("\nEnter Student ID: ", end='')
ID = int(input())
Name = ""
for i in range(45):
    row_number = 3+i
    Id = int(sheet.cell_value(row_number, 2))
    if(Id == ID):
        Name = sheet.cell_value(row_number, 3)
        break
    
if(Name == ""):
    print("\nID not found.\n\n\n")
    exit()
    
print("\nStudent Name is: " + Name + "\n\n")
add_text_in_file(ID, Name)


