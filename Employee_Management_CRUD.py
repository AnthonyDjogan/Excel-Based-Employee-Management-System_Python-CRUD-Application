import pandas as pd
from tabulate import tabulate


# Global Variable and Function
temporary_value = {         # This variable is used for storing our input when using menu create and update
    'EMPLOYEE ID': [], 
    'RESIDENCE': [],
    'NAME': [],
    'GENDER': [],
    'DEPARTMENT': []
}

def Call_menu(*args, menu_title):
    print('----------------------------')
    print("{:^29}".format(menu_title))
    print('----------------------------')
    for i, option in enumerate(args, start=1):
        print(f"{i}. {option}")
    print('----------------------------')
    input_menu = input('Select Menu: ').strip()
    print()
    return input_menu

def Menu_error_message(*args):
    print('Menu', *args, 'does not exist\n')

def Input_error_message(*args):
    len_char = len(*args)
    print('+' * len_char)
    print(*args)
    print('+' * len_char + '\n')

def Input_checker(input, column, dept):
    if column == 'RESIDENCE':  
        check_character_RESIDENCE = ''.join(input.split())   
        if check_character_RESIDENCE.isalpha():  
            return True
        else:
            print('RESIDENCE can only be filled with letters')
            return False
        
    elif column == 'NAME':
        if input.isspace():     
            print('Name cannot be left blank. Please enter a valid name.')

            return False
        else:
            return True
        
    elif column == 'GENDER':
        if input not in ['M', 'F']:
            print("Invalid input. Please enter 'M' for male or 'F' for female")
            return False
        else:
            return True
        
    elif column == 'DEPARTMENT':
        if input in dept.values:
            return True
        else:
            print('Please enter a valid department name')
            return False
        
def ID_generator(data, used_id, department, gender, residence, name):
    global temporary_value
    DEPARTMENT_initial = ''.join([word[0].upper() for word in department.split()])      
                                                                                        
    residence_initial = ''.join([word[0].upper() for word in residence.split()])   
    name_length = len(name.strip())
    id_generator = f'{DEPARTMENT_initial}{gender}{residence_initial}{name_length}'

    all_ids = data['EMPLOYEE ID'].tolist() + used_id['EMPLOYEE ID'].tolist()           
    while id_generator in all_ids:
        name_length += 1
        id_generator = f'{DEPARTMENT_initial}{gender}{residence_initial}{name_length}'
    
    temporary_value.update({'EMPLOYEE ID': id_generator})
    return id_generator

# Menu Read
def Show_specific_data(data): 
    ask_input_ID = input('Input EMPLOYEE ID: ').upper().strip()

    if ask_input_ID in data["EMPLOYEE ID"].values:
        print(tabulate(data.loc[data['EMPLOYEE ID'] == ask_input_ID].values, headers= data.columns, tablefmt='fancy_grid'))
    else:
        print('Employee ID not found')
    
    while True:
        find_other_ID = input('Search other ID [Y/N]:').upper().strip()
        print()
        if find_other_ID == 'Y':
            Show_specific_data(data)
        elif find_other_ID == 'N':
            break
        else:
            Input_error_message('Invalid Input. Input "Y" to search other ID or "N" to exit')

# Menu Create
def Input_data_employee(data, used_id, dept, value):
    global employee_path, usedID_path
    while True:
        column_input_RESIDENCE = 'RESIDENCE'      
        input_RESIDENCE = input('RESIDENCE: ').title().strip()
       
        if Input_checker(input_RESIDENCE, column_input_RESIDENCE, dept= dept) == True:
            value.update({'RESIDENCE': input_RESIDENCE})
            break
        else:
            continue

    while True:
        column_input_NAME = 'NAME'
        input_NAME = input('NAME: ').title()   
                                                        
        if Input_checker(input_NAME, column_input_NAME, dept= dept) == True:
            value.update({'NAME': input_NAME})
            break
        else:
            continue

    while True:
        column_input_GENDER = 'GENDER'
        input_GENDER = input('GENDER [M/F]: ').upper().strip()

        if Input_checker(input_GENDER, column_input_GENDER, dept= dept) == True:
            value.update({'GENDER': input_GENDER})
            break
        else:
            continue

    while True:
        column_input_DEPARTMENT = 'DEPARTMENT'
        print(tabulate(dept, headers='keys', tablefmt='fancy_grid'))

        input_DEPARTMENT = input('DEPARTMENT: ').title().strip()
        if Input_checker(input_DEPARTMENT, column_input_DEPARTMENT, dept= dept) == True:
            value.update({'DEPARTMENT': input_DEPARTMENT})
            break
        else:
            continue  
    generated_id = ID_generator(data= data, used_id= used_id, department= input_DEPARTMENT, gender= input_GENDER, residence= input_RESIDENCE, name= input_NAME)
    
    while True:
        input_confirmation_add_data = input(f"Are you sure you want to add data with Employee ID {value['EMPLOYEE ID']}? [Y/N]: ").upper().strip()
        print()
        if input_confirmation_add_data == 'Y':
            data = pd.concat([data, pd.DataFrame.from_dict(value, orient= 'index').T], ignore_index= True)
            
            print('\nData Saved')
            data.to_excel(employee_path, index=False)

            used_id.loc[len(used_id)] = generated_id
            used_id.to_excel(usedID_path, index= False)
            print(tabulate(data, headers='keys', tablefmt='fancy_grid'))
            break
        elif input_confirmation_add_data == 'N':
            break
        else:
            Input_error_message('Invalid input. Input [Y] to proceed or [N] to cancel')

# Menu Update
def Update_data(data, index_change, column_change, value): 
    global employee_path  
    while True:
        input_konfirmasi_update_data = input(f"Are you sure you want to update data for ID {data.loc[index_change]['EMPLOYEE ID'].values}? [Y/N]: ").upper().strip()
        if input_konfirmasi_update_data == 'Y':
            data.loc[index_change, column_change] = value      
            data.to_excel(employee_path, index=False)
            print(tabulate(data.loc[index_change, :], headers='keys', tablefmt='fancy_grid'))
            break
       
        elif input_konfirmasi_update_data == 'N':
            print('\nThe update process has been canceled.\n')
            break

        else:
            Input_error_message('Invalid Input. Input [Y] to proceed or [N] to cancel')

def Input_data_to_change(data, dept):
    while True:     
        input_id_update = input('EMPLOYEE ID: ').upper().strip()   

        if input_id_update in data['EMPLOYEE ID'].values:
            print(tabulate(data.loc[data['EMPLOYEE ID'] == input_id_update].values, headers= data.columns, tablefmt='fancy_grid'))
            index_to_update = data.loc[data['EMPLOYEE ID'] == input_id_update].index
        
        else:
            Input_error_message(f'There is no Employee with ID: {input_id_update}')
            break
  
        column_to_update = None                   
        while True:                
            input_confirmation_change_data = input('Input [Y] to proceed or [N] to cancel: ').upper().strip()
    
            if input_confirmation_change_data == 'Y':
                while True:
                    input_column_change = input('Please enter the column name: ').upper().strip()
                    
                    if input_column_change in data.columns:      
                        column_to_update = input_column_change
                        break

                    else:
                        Input_error_message(f'Column {input_column_change} does not exist')
                        continue
                          
                if column_to_update == 'RESIDENCE':
                    while True:
                        input_new_RESIDENCE = input(f'Input new {input_column_change}: ').title().strip()

                        if Input_checker(input_new_RESIDENCE, column_to_update, dept= dept) == True:
                            Update_data(data= data, index_change= index_to_update, column_change= column_to_update, value= input_new_RESIDENCE)
                            return
                        else:
                            continue

                elif column_to_update == 'NAME':
                    while True:
                        input_new_NAME = input(f'Input new {input_column_change}: ').title()

                        if Input_checker(input_new_NAME, column_to_update, dept= dept) == True:
                            Update_data(data= data, index_change= index_to_update, column_change= column_to_update, value= input_new_NAME.strip())
                            return
                        else:
                            continue

                elif column_to_update == 'GENDER':
                    while True:
                        input_new_GENDER = input(f'Input new {input_column_change}: ').upper().strip()

                        if Input_checker(input_new_GENDER, column_to_update, dept= dept) == True:
                            Update_data(data= data, index_change= index_to_update, column_change= column_to_update, value= input_new_GENDER)
                            return
                        else:
                            continue

                elif column_to_update == 'DEPARTMENT':
                    while True:
                        print(tabulate(dept, headers='keys', tablefmt='fancy_grid'))

                        input_new_DEPARTMENT = input(f'Input new {input_column_change}: ').title().strip()

                        if Input_checker(input_new_DEPARTMENT, column_to_update, dept= dept) == True:
                            Update_data(data= data, index_change= index_to_update, column_change= column_to_update, value= input_new_DEPARTMENT)
                            return
                        else:
                            continue

                elif input_column_change == 'EMPLOYEE ID':
                    Input_error_message('Can not change EMPLOYEE ID')
                    return
            
            elif input_confirmation_change_data == 'N':
                return
            else:
                Input_error_message('Invalid input. Input [Y] to proceed or [N] to cancel')

# Menu Delete
def Delete_data(data, dept):
    global employee_path
    input_id_delete = input('EMPLOYEE ID: ').upper().strip()       
    
    if input_id_delete not in data['EMPLOYEE ID'].values:
        Input_error_message(f'{input_id_delete} does not exist')
        return
    
    index_to_delete = data.loc[data['EMPLOYEE ID'] == input_id_delete].index         
    print(tabulate(data.loc[data['EMPLOYEE ID'] == input_id_delete], headers= data.columns, tablefmt='fancy_grid'))

    while True:
        input_confirmation_delete_data = input(f"Are you sure you want to delete data with ID {data.loc[index_to_delete, 'EMPLOYEE ID'].values}? [Y/N]: ").upper().strip()
        if input_confirmation_delete_data == 'Y':
            data = data.drop(index_to_delete)
            data.to_excel(employee_path, index=False)
            print('\nData has been successfully deleted\n')
            break

        elif input_confirmation_delete_data == 'N':
            print('\nThe deletion process has been canceled\n')
            break
        
        else:
            Input_error_message('Invalid input. Input [Y] to proceed or [N] to cancel')
            continue

def Read_file(path, sheet_name):
    pd.read_excel(path, sheet_name, dtype= 'object')

def Save_file(path, sheet_name):
    pd.to_excel(path, sheet_name, index= False)

def Main_flow():
    while True:
        global employee_path, dept_path, usedID_path
        # Define DF
        employee_path = r"C:\Users\msi pc\Desktop\Data Science\Purwadhika\KELAS DS\Solo\Employee Data.xlsx"
        dept_path = r"C:\Users\msi pc\Desktop\Data Science\Purwadhika\KELAS DS\Solo\Available Department.xlsx"
        usedID_path = r'C:\Users\msi pc\Desktop\Data Science\Purwadhika\KELAS DS\Solo\Used ID.xlsx'

        employee_data = pd.read_excel(employee_path, dtype= 'object')
        available_department = pd.read_excel(dept_path, dtype= 'object')
        used_ID = pd.read_excel(usedID_path, dtype= 'object')

        main_menu_input = Call_menu('Employee Data', 'Add Data', 'Change Data', 'Delete Data', 'Exit', menu_title= 'Main Menu')
        if main_menu_input == '1':      # Menu Read
            while True:
                sub_menu_1_input = Call_menu('Show All Data', 'Search ID', 'Main Menu', menu_title= 'Employee Data')
                if sub_menu_1_input == '1':     # Show all employee data
                    print(tabulate(employee_data, headers='keys', tablefmt='fancy_grid'))
                elif sub_menu_1_input == '2':   # Show specific data based on ID
                    Show_specific_data(employee_data)
                elif sub_menu_1_input == '3':   # Return to main menu
                    break
                else:
                    Menu_error_message(sub_menu_1_input)

        elif main_menu_input == '2':    # Menu Create
            while True:
                sub_menu_2_input = Call_menu('Add Employee Data', 'Main Menu', menu_title= 'Add Data')
                if sub_menu_2_input == '1':     # Add new employee
                    Input_data_employee(data= employee_data, used_id= used_ID, dept= available_department, value= temporary_value)
                elif sub_menu_2_input == '2':   # Return to main menu
                    break
                else:
                    Menu_error_message(sub_menu_2_input)

        elif main_menu_input == '3':    # Menu Update
            while True:
                sub_menu_3_input = Call_menu('Change Employee Data', 'Main Menu', menu_title= 'Change Data')
                if sub_menu_3_input == '1':     # Change employee data based on ID
                    Input_data_to_change(data= employee_data, dept= available_department)
                elif sub_menu_3_input == '2':   # Return to main menu
                    break
                else:
                    Menu_error_message(sub_menu_3_input)

        elif main_menu_input == '4':    # Menu Delete
            while True:
                sub_menu_4_input = Call_menu('Delete Employee Data', 'Main Menu', menu_title= 'Delete Data')
                if sub_menu_4_input == '1':     # Delete employee data based on ID
                    Delete_data(data= employee_data, dept= available_department)
                elif sub_menu_4_input == '2':   # Return to main menu
                    break
                else:
                    Menu_error_message(sub_menu_4_input)

        elif main_menu_input == '5':    
            exit()
        else:
            Menu_error_message(main_menu_input)

Main_flow()