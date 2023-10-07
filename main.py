import xlrd
from sympy import Matrix
workbook = xlrd.open_workbook('ATTENDANCE_SHEET.xls')
sheet = workbook.sheet_by_name('DynamicReport')  # Replace 'Sheet1' with the actual sheet name

def add_text_in_file(ID, Name):
    try:
        with open('origin.txt', 'r') as file:
            # Read the content of the file
            file_content = file.read()
        
        
        
        
        print(Name)
        # Replace the old text with the new text
        file_content = file_content.replace('@Name@', Name)
        file_content = file_content.replace('@ID@', str(ID))
        d1 = int((ID%1000)/100)
        d2 = int((ID%100)/10)
        d3 = (ID%10)
        
        
        expressions = [
            (3*d1 + 6*d2 - 4*d3) % 10 + d1,
            (7*d1 - 2*d2 + 8*d3) % 10 + d2,
            -((5*d1 + 9*d2 - 1*d3) % 10 + d3+1),
            (2*d1 - 7*d2 + 5*d3) % 10 + d2,
            (9*d1 + 1*d2 + 4*d3) % 10 + d1,

            (1*d1 - 4*d2 + 7*d3) % 10 + d3,
            -((8*d1 + 5*d2 - 2*d3) % 10 + d1+1),
            (6*d1 - 3*d2 + 9*d3) % 10 + d2,
            (4*d1 + 2*d2 + 8*d3) % 10 + d3,
            (5*d1 - 6*d2 + 7*d3) % 10 + d1,

            -((9*d1 + 3*d2 - 2*d3) % 10 + d2+1),
            (2*d1 + 7*d2 + 1*d3) % 10 + d1,
            -((4*d1 - 5*d2 + 6*d3) % 10 + d3+1),
            -((6*d1 + 9*d2 + 3*d3) % 10 + d2+1),
            (8*d1 - 1*d2 + 7*d3) % 10 + d1,

            (7*d1 + 5*d2 - 8*d3) % 10 + d2,
            -((4*d1 - 2*d2 + 9*d3) % 10 + d1+1),
            -((6*d1 + 8*d2 - 3*d3) % 10 + d3+1),
            -((3*d1 + 1*d2 + 7*d3) % 10 + d2+1),
            -((5*d1 - 9*d2 + 4*d3) % 10 + d3+1)
        ]

        # Reshape the expressions into a 4x5 matrix
        M_values = [expressions[i:i+5] for i in range(0, len(expressions), 5)]
        M_a = Matrix(M_values)
        
        for k in range(4): 
            M_a[4+5*k] = d1*M_a[0+5*k] + d2*M_a[1+5*k] - d3*M_a[2+5*k] + (d1-d3)*M_a[3+5*k]
        
        if(d3==0 or d3==5 or d3==9):
            for k in range(5):
                M_a[5+k] = d1*M_a[0+k] + d2*M_a[10+k] - d3*M_a[15+k]
            if(d3==0):
                M_a[9] = d1*M_a[4] + d2*M_a[14] - d3*M_a[19]+1
        
        
        file_content = file_content.replace('@A11@', str(M_a[0]))
        file_content = file_content.replace('@A12@', str(M_a[1]))
        file_content = file_content.replace('@A13@', str(M_a[2]))
        file_content = file_content.replace('@A14@', str(M_a[3]))
        file_content = file_content.replace('@A15@', str(M_a[4]))

        file_content = file_content.replace('@A21@', str(M_a[5]))
        file_content = file_content.replace('@A22@', str(M_a[6]))
        file_content = file_content.replace('@A23@', str(M_a[7]))
        file_content = file_content.replace('@A24@', str(M_a[8]))
        file_content = file_content.replace('@A25@', str(M_a[9]))

        file_content = file_content.replace('@A31@', str(M_a[10]))
        file_content = file_content.replace('@A32@', str(M_a[11]))
        file_content = file_content.replace('@A33@', str(M_a[12]))
        file_content = file_content.replace('@A34@', str(M_a[13]))
        file_content = file_content.replace('@A35@', str(M_a[14]))

        file_content = file_content.replace('@A41@', str(M_a[15]))
        file_content = file_content.replace('@A42@', str(M_a[16]))
        file_content = file_content.replace('@A43@', str(M_a[17]))
        file_content = file_content.replace('@A44@', str(M_a[18]))
        file_content = file_content.replace('@A45@', str(M_a[19]))
        
        
        
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
        M_values = [expressions[j:j+6] for j in range(0, len(expressions), 6)]
        M_b = Matrix(M_values)
        
        for k in range(4): 
            M_b[5+6*k] = (d1+1)*M_b[0+6*k] + d2*M_b[1+6*k] - d3*d3*M_b[2+6*k] + (d1-d3)*M_b[3+6*k] + (d1+ d1*d2)*M_b[4+6*k]
        
        
        if(d3==0 or d3==1 or d3==9):
            for k in range(6):
                M_b[6+k] = d1*M_b[0+k] + d2*M_b[12+k] - d3*M_b[18+k]
        if(d3==0):
            for k in range(6):
                M_b[18+k] = (d3+1)*M_b[0+k] + d2*M_b[12+k] - d1*M_b[6+k]
        

        file_content = file_content.replace('@B11@', str(M_b[0]))
        file_content = file_content.replace('@B12@', str(M_b[1]))
        file_content = file_content.replace('@B13@', str(M_b[2]))
        file_content = file_content.replace('@B14@', str(M_b[3]))
        file_content = file_content.replace('@B15@', str(M_b[4]))
        file_content = file_content.replace('@B16@', str(M_b[5]))

        file_content = file_content.replace('@B21@', str(M_b[6]))
        file_content = file_content.replace('@B22@', str(M_b[7]))
        file_content = file_content.replace('@B23@', str(M_b[8]))
        file_content = file_content.replace('@B24@', str(M_b[9]))
        file_content = file_content.replace('@B25@', str(M_b[10]))
        file_content = file_content.replace('@B26@', str(M_b[11]))

        file_content = file_content.replace('@B31@', str(M_b[12]))
        file_content = file_content.replace('@B32@', str(M_b[13]))
        file_content = file_content.replace('@B33@', str(M_b[14]))
        file_content = file_content.replace('@B34@', str(M_b[15]))
        file_content = file_content.replace('@B35@', str(M_b[16]))
        file_content = file_content.replace('@B36@', str(M_b[17]))

        file_content = file_content.replace('@B41@', str(M_b[18]))
        file_content = file_content.replace('@B42@', str(M_b[19]))
        file_content = file_content.replace('@B43@', str(M_b[20]))
        file_content = file_content.replace('@B44@', str(M_b[21]))
        file_content = file_content.replace('@B45@', str(M_b[22]))
        file_content = file_content.replace('@B46@', str(M_b[23]))







        # Write the updated content back to the file
        with open('file.txt', 'a') as file:
            file.write(file_content)

        print('Text replacement successful.')
    except FileNotFoundError:
        print(f'The file "{file_path}" was not found.')
    except Exception as e:
        print(f'An error occurred: {str(e)}')




for i in range(45):
    row_number = 3+i
    ID = int(sheet.cell_value(row_number, 2))
    Name = sheet.cell_value(row_number, 3)
    add_text_in_file(ID, Name)


