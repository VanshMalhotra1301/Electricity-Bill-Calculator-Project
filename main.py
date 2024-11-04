import openpyxl
from datetime import datetime

def get_appliance_data():
    appliances = []
    while True:
        name = input("Enter appliance name (or 'done' to finish): ")
        if name.lower() == 'done':
            break
        power = float(input("Enter power consumption in watts (W): "))
        hours = float(input("Enter hours used per day: "))
        quantity = int(input("Enter number of appliances: "))
        
        appliance = {
            'name': name,
            'power': power,
            'hours': hours,
            'quantity': quantity
        }
        appliances.append(appliance)
    
    return appliances

def calculate_costs(appliances, cost_per_kwh):
    results = []
    total_cost = 0  # Initialize total cost
    for appliance in appliances:
        # Calculate energy consumed in kWh
        energy_consumed = (appliance['power'] / 1000) * appliance['hours'] * appliance['quantity']
        
        # Calculate the cost for this appliance in rupees
        cost = energy_consumed * cost_per_kwh
        
        # Appending results with a dictionary
        results.append({
            'name': appliance['name'],
            'power': appliance['power'],
            'hours': appliance['hours'],
            'quantity': appliance['quantity'],
            'energy_consumed': energy_consumed,
            'cost': cost
        })
        
        total_cost += cost  # Add to total cost
        monthly_cost = total_cost*30
    
    return results, monthly_cost  # Return total cost as well

def export_to_excel(results, total_cost):
    # Generate a unique filename with a timestamp
    filename = f'electricity_costs_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    # Create a new workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Electricity Costs"
    
    # Write the header
    headers = ['Appliance Name', 'Power (W)', 'Hours/Day', 'Quantity', 'Energy Consumed (kWh)', 'Cost (INR)']
    sheet.append(headers)
    
    # Write the data
    for result in results:
        row = [
            result['name'],
            result['power'],
            result['hours'],
            result['quantity'],
            result['energy_consumed'],
            result['cost']
        ]
        sheet.append(row)
    
    # Write total cost at the bottom
    sheet.append([])
    sheet.append(['Total Monthly Cost (INR)', total_cost])
    
    # Save the workbook
    workbook.save(filename)
    print(f"Data exported to {filename} successfully!")

def main():
    print("Welcome to the Electricity Bill Calculator!")
    
    # Step 1: Get appliance data from the user
    appliances = get_appliance_data()
    
    # Step 2: Get cost per kWh in rupees
    cost_per_kwh = float(input("Enter the cost of electricity per kWh in your location (in INR): "))
    
    # Step 3: Calculate costs for each appliance
    results, total_cost = calculate_costs(appliances, cost_per_kwh)
    
    # Step 4: Export results to Excel
    export_to_excel(results, total_cost)

# Run the program
if __name__ == "__main__":
    main()
