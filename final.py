try:
    import xlwt
    from xlwt import Workbook
    spreadsheet = True
except:
    spreadsheet = False
    print("No spreadsheet support.")
    print("Fix with `pip install xlwt`")

spreadsheet_order = ["name", "id", "wage", "hours", "gross", "tax", "post", "dependents"]

employees = {}

tax = {"state": 0.056, 
     "federal": 0.079}

def delete_employee():
    id = input("Employee's ID: ")
    if not id in employees:
        print("No employee to delete")
        print()
        return
    print(f"Are you sure you wish to delete {employees[id]["name"]["last"]}, {employees[id]["name"]["first"]}?")
    choice = input("y/n: ")
    if choice.lower() == "y":
        employees.pop(id, None)
        print("Employee record deleted.")
    else:
        print("Cancelling.")
    print()
    return


def add_employee():
    id = input("Employee's ID (NOT NAME): ")
    if id in employees:
        print("ID exists. Please edit employees instead!")
        print()
        return
    
    first = input("Employee's first name: ")
    last = input("Employee's last name: ")
    dependents = input ("Number of dependents: ")
    try:
        int(dependents)
    except:
        "Not a number: Returning 0."
        dependents = 0
    hours = input("Hours worked: ")
    try:
        float(hours)
    except:
        "Not a number: Returning 0."
        hours = 0
    wage = input("Wage: ")
    try:
        float(wage)
    except:
        "Not a number: Returning 0."
        wage = 0

    employees[id] = {
        "name":{
            "first": first,
            "last": last
        },
        "dependents":int(dependents),
        "hours":float(hours),
        "wage": round(float(wage), 2),
        "id":id
    }
    print(employees[id]["wage"])
    print("Thank you!")
    print()

def edit_employee(i:int = 0, id = None):
    if id == None:
        id = input("Which Employee would you edit: ")

    if not str(id) in employees:
        print("Error: No ID found.")
        if i < 3:
            print()
            edit_employee(i + 1)
        else:
            print("Returning")
            print()
        return

    print("What would you like to edit:")
    print(f"1. Hours worked: {employees[id]["hours"]}")
    print(f"2. Dependents: {employees[id]["dependents"]}")
    print(f"3. Last Name: {employees[id]["name"]["last"]}")
    print(f"4. First name: {employees[id]["name"]["first"]}")
    print(f"5. Wage: {employees[id]["wage"]}")
    edit = input()

    change = input("New value: ")

    try:
        match int(edit): # I would usually avoid repeating actions, but due to the names it's easier to just do this. 
            case 1:
                employees[id]["hours"] = int(change)
            case 2:
                employees[id]["dependents"] = int(change)
            case 3:
                employees[id]["name"]["first"] = change
            case 4:
                employees[id]["name"]["last"] = change
            case 5:
                employees[id]["wage"] = round(float(change), 2)
            case _:
                print("To edit ID, make a new user.") # I forsee a ton of issues with allowing people to edit IDs
                print("No value selected: no changes made.")
                print()
                return
        print("Value changed!")
        print("Change another value? y/n")
        answer = input()
        if answer.lower() == "y":
            edit_employee(id=id)
        return
    except:
        print("Error changing value.")


def render_dict(id:str):
    try:
        print(f"Employee {employees[id]["name"]["last"]}, {employees[id]["name"]["first"]}:")
        print(f"ID: {employees[id]["id"]}")
        print(f"Wage: {employees[id]["wage"]}")
        print(f"Working hours: {employees[id]["hours"]}")
        print(f"Dependencies: {employees[id]["dependents"]}")

    except:
        print("Error showing list.")
        debug = input("Debug? y/n: ")
        if debug.lower() == "y":
            print(employees[id])
    print()
    return

def view_employees(id = True):
    if not type(id) is bool:
        if not str(id) in employees:
            print(f"Error showing {id}")
            return
        render_dict(str(id))
        return
    else:
        for x in employees:
            render_dict(x)

def gross_pay(hourly:float, wage:float):
    if hourly > 40:
        overtime = hourly - 40
        total = (wage * 40)
        total += (wage * overtime * 1.5)
        return total
    if hourly > 0:
        return hourly * wage
    print("Error: bad hours.")
    print(f"Please fix {hourly} to be over 0")
    return 0

def post_tax(income):
    total_tax = 0
    for percent in tax.values():
        total_tax += float(percent) * float(income)
    return total_tax

        


def Calculate_pay(id:str):
    if not id in employees:
        print("Error: No ID found.")
        print()
        return
    print(f"{employees[id]["name"]["last"]}, {employees[id]["name"]["last"]}")
    try:
        wage = employees[id]["wage"]
        hourly = employees[id]["hours"]
        gross = gross_pay(hourly, wage)
        print(f"Gross pay: {gross}")
        tax = post_tax(gross)
        print(f"Taxes: {tax}")
        employees[id]["gross"] = gross
        employees[id]["tax"] = tax
        employees[id]["post"] = gross - tax
        print(f"Post: {gross - tax}")
        print()
    except:
        print("Unable to process. Please edit employee.")



def spreadthemsheets():
    wb = Workbook()
    sheet = wb.add_sheet("Employees")

    sheet.write(0, 0, "Employees")
    for i, name in enumerate(spreadsheet_order):
        if name != "name":
            sheet.write(0, i, name.capitalize())
    i = 0
    for employee in employees:
        value = employees[employee]
        Calculate_pay(employee)
        i += 1
        sheet.write(i, 0, f"{value["name"]["last"]}, {value["name"]["first"]}")
        for stat in value:
            if stat != "name":
                place = spreadsheet_order.index(stat)
                sheet.write(i, place, str(value[stat]))
    wb.save("Employees sheet.xls")




def fill_with_fake_values():
    print("Values added")
    print("ID's 1 through 4 have been overwrited.")
    print()

    id = 1
    employees[id] = {
        "name":{
            "first": "Water",
            "last": "Boy"
        },
        "dependents":int(0),
        "hours":float(20),
        "wage": round(float(14.999), 2),
        "id":id
    }
    
    id = 2
    employees[id] = {
        "name":{
            "first": "Lava",
            "last": "Girl"
        },
        "dependents":int(0),
        "hours":float(21.25),
        "wage": round(float(14), 2),
        "id":id
    }
    
    id = 3
    employees[id] = {
        "name":{
            "first": "Lazy",
            "last": "Bum"
        },
        "dependents":int(0),
        "hours":float(5),
        "wage": round(float(25), 2),
        "id":id
    }
    
    id = 4
    employees[id] = {
        "name":{
            "first": "Triangle",
            "last": "Gangle"
        },
        "dependents":int(2),
        "hours":float(15),
        "wage": round(float(16.20), 2),
        "id":id
    }




def action():
    print("What would you like to do?")
    print("1. Show Employees")
    print("2. Calculate pay")
    if spreadsheet:
        print("3. Create Spreadsheet")
    else:
        print("- Spreadsheet not avaliable -")
    print("4. Edit Employee")
    print("5. Add Employee")
    print("6. Remove Employee")
    print("7. Debug Test Employees")
    print("0. Exit")
    choice = input()

    match int(choice):
        case 1:
            view_employees()
        case 2:
            for x in employees:
                Calculate_pay(x)
        case 3:
            if spreadsheet:
                try:
                    spreadthemsheets()
                except:
                    print("Error: Sheet probably open.")
                    print()
        case 4:
            edit_employee()
        case 5: 
            add_employee()
        case 6:
            delete_employee()
        case 7:
            fill_with_fake_values()
        case 0:
            return False
    return True

loop = True
while loop:
    loop = action()



