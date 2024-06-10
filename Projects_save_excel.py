import time
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime

wb = Workbook("Test.xlsx")
path_file = ("Test.xlsx")

def new_project():
    id_input = input("Enter the id of the proyect")
    name_input = input("Enter the name of the project")
    status_input = input("Enter the status of the project")
    delivery_input = input("Enter the delivery date of the project ")
    start_time = time.asctime()
    current_data2 = [id_input,
                     name_input,
                     status_input,
                     delivery_input,
                     start_time]
    print("[✓] Proyect Created.", "\n", "ID:", id_input, "\n"
          "Name:", name_input, "\n"
          "Status:", status_input, "\n"
          "Delivery date:", delivery_input, "\n"
          "Start time:", start_time, "\n")
    return current_data2


def save_progress(current_data2):
    wb = load_workbook("Test.xlsx")
    sheet = wb["Projects"]
    id = current_data2[0]
    name = current_data2[1]
    status = current_data2[2]
    delivery_date = current_data2[3]
    start_time = current_data2[4]
    datos = [id, name, status, delivery_date, start_time]
    sheet.append(datos)
    wb.save(path_file)


def finish_projects():
    wb = load_workbook("Test.xlsx")
    sheet = wb["Projects"]
    id_input = input("Enter ID project...")
    id = (int(id_input) + 1)
    finish_project_input = input("Are you sure you want to finalize project " + str(sheet.cell(row = id, column = 2).value) + "? y/n ").strip().lower()
    if finish_project_input == "y":
        sheet.cell(row = id, column = 3).value = "Terminado"
        print(sheet.cell(row = id, column = 3).value)
        horario_de_projecto_finalizado = time.asctime()
        sheet.cell(row = id, column = 6).value = horario_de_projecto_finalizado
        horario_excel = sheet.cell(id, column = 5).value
        horario_final = horario_de_projecto_finalizado
        horario_inicio = datetime.strptime(horario_excel,"%a %b %d %H:%M:%S %Y")
        horario_actual = datetime.strptime(horario_final, "%a %b %d %H:%M:%S %Y")
        duracion = horario_actual - horario_inicio
        duracion_en_horas = duracion.total_seconds() / 3600
        project_name = sheet.cell(id, column = 2).value
        print(f"Congratulations, the project {project_name} lasted {round(duracion_en_horas, 1)} hours.")
        sheet.cell(row = id, column = 7). value = round(duracion_en_horas, 1)
        wb.save("Test.xlsx")
    elif finish_project_input == "n":
        print("Back to the menu!")
        return main()
            
def read_projects():
    wb = load_workbook("Test.xlsx")
    sheet = wb["Projects"]
    for row in sheet.iter_rows():
        line = []
        for cell in row:
            line.append(str(cell.value))
        print("\t".join(line))


def edit():
    wb = load_workbook("Test.xlsx")
    sheet = wb["Projects"]
    id_input = input("Enter ID project...")
    id_calculator= (int(id_input) + 1)
    id = (id_calculator)
    edit_project_input = input("""What do you want to modify?
        1.Name project
        2.Status
        3.Delivery date.
        """)
    if edit_project_input == "1":
        edit_name_input = input("""What do you want to change it to?""")
        sheet.cell(row = id, column = 2).value = edit_name_input
        print(sheet.cell(row = id, column = 2).value)
        save_edit = input("Do you want save it? y/n").strip().lower()
        if save_edit == "y":
            wb.save("Test.xlsx")
        if save_edit == "n":
            return main()
        
    elif edit_project_input == "2":
        edit_status_input = input("""What do you want to change it to?""")
        sheet.cell(row = id, column = 3).value = edit_status_input
        print(sheet.cell(row = id, column = 3).value)
        save_edit = input("Do you want save it? y/n").strip().lower()
        if save_edit == "y":
            wb.save("Test.xlsx")
        if save_edit == "n":
            return main()
        
    elif edit_project_input == "3":
        edit_delivery_input = input("""What do you want to change it to?""")
        sheet.cell(row = id, column = 4).value = edit_delivery_input
        print(sheet.cell(row = id, column = 4).value)
        save_edit = input("Do you want save it? y/n").strip().lower()
        if save_edit == "y":
            wb.save("Test.xlsx")
        if save_edit == "n":
            return main()

def main():
    print("Welcome to Excel Simple Work Data")
    print("\nYour projects are:")

    while True:
        option = input("""
          +==========================================+
          |                                          |
          |               Work Project               |
          |                                          |
          +==========================================+
        
          [?] select an option:
        
          1. New project
          2. Search for a proyect
          3. Finish project
          4. Exit
          0. Options
          
        > """)
        if option == "1":
            new_project()
            save_it = input("Do you want save it? y/n").strip().lower()
            if save_it == "y":
                save_progress()

        if option == "2":
            print("""The proyects are...
            \t""")
            read_projects()
            print("\n select an option:")
            option_input = input("""
            1. Edit project.
            2. Delete.
            3. Return main.
            """)
            if option_input == "1":
                edit()
            if option_input == "2":
                delete_input = input("Enter ID project...")
                delete_calulator = (int(delete_input) + 1)
                delete = input("Are you sure to delete it? y/n").strip().lower()
                if delete == "y":
                    wb = load_workbook("Test.xlsx")
                    sheet = wb["Projects"]
                    sheet.delete_rows(delete_calulator)
                    save_input = input("Do you want save it? y/n").strip().lower()
                    if save_input == "y":
                        wb.save("Test.xlsx")

        if option == "3":
            finish_projects()
            return main()

        if option == "4":
            print("Have a good day!")
            return False

        if option == "0":
            create_workbook_input = input("Create a new archive? y/n").lower()
            if create_workbook_input == "y":
                file_name_input = input("Perfect! What name do you want for the file?")
                confirm_workbook_input = input("Are you sure you want to give it this name " + file_name_input + "? y/n").lower()
                if confirm_workbook_input == "y":
                    create_workbook=(file_name_input)
                    wb = Workbook()
                    save_workbook = input("Do you want save it? y/n")
                    if save_workbook == "y":
                        print("Perfect! the file " + create_workbook + " has been saved")
                        hoja = wb.active
                        hoja.title = "Projects"
                        sheet = wb["Projects"]
                        columnas = ["ID", "Name", "Status", "Delivery date", "Start Time, Finish time, Project duration hours"]
                        sheet.append(columnas)
                        wb.save(create_workbook + ".xlsx")
                        return main()

if __name__ == "__main__":
    main()