import pandas as pd
from datetime import datetime, timedelta

officer_dict = {}

# Function to parse unavailability dates
def parse_unavailability_dates(date_str, year, month):
    dates = []
    if '-' in date_str:
        start_day, end_day = map(int, date_str.split('-'))
        start_date = datetime(year, month, start_day)
        end_date = datetime(year, month, end_day)
        current_date = start_date
        while current_date <= end_date:
            dates.append(current_date)
            current_date += timedelta(days=1)
    else:
        day = int(date_str)
        dates.append(datetime(year, month, day))
    return dates

# Read input Excel file
input_file = 'officers_unavailability.xlsx'
df = pd.read_excel(input_file)

#creates the dictionary of names and sets their initial duty count to 0
officer_dict = {name: 0 for name in df['Name'].unique()}

# Extracting the current and next month
today = datetime.today()
first_day_next_month = (today.replace(day=1) + timedelta(days=32)).replace(day=1)
last_day_next_month = (first_day_next_month.replace(month=first_day_next_month.month % 12 + 1) - timedelta(days=1))

# Generate a list of all dates in the next month
next_month_dates = pd.date_range(start=first_day_next_month, end=last_day_next_month).to_pydatetime().tolist()

# Parse unavailability dates for each person
unavailability_dict = {}
for idx, row in df.iterrows():
    #print(row.Unavailability)
    print(row.Name)
    name = row['Name']
    unavailability_str = str(row['Unavailability'])
    unavailability_dates = []
    if pd.isna(unavailability_str) or unavailability_str != 'nan':
        for date_part in unavailability_str.split(','):
            unavailability_dates.extend(parse_unavailability_dates(date_part.strip(), first_day_next_month.year, first_day_next_month.month))
    unavailability_dict[name] = unavailability_dates

# Create a dictionary to track duty assignments
duty_schedule = {date: None for date in next_month_dates}

# Function to find the next available person
def get_next_available_person(date):
    best_candidate = None
    for idx, row in df.iterrows():
        name = row['Name']
        #print(unavailability_dict[name][0].day)
        #print(date.day)
        unavailable_days = [d.day for d in unavailability_dict[name]]
        if date.day not in unavailable_days:

            #if all(duty_schedule[d] != name for d in duty_schedule if duty_schedule[d] is not None):
                if officer_dict[name] == 0:
                    officer_dict[name] += 1
                    return name
                
                if best_candidate is None or officer_dict[name] < officer_dict[best_candidate]:
                    best_candidate = name
    
    if best_candidate:
        officer_dict[best_candidate] += 1
                
    return best_candidate

# Assign duties
for date in next_month_dates:
    person = get_next_available_person(date)
    if person:
        duty_schedule[date] = person

# Convert the duty schedule to a DataFrame
duty_list_df = pd.DataFrame(list(duty_schedule.items()), columns=['Date', 'Assigned Person'])

# Save the duty list to an Excel file
output_file = 'duty_list.xlsx'
duty_list_df.to_excel(output_file, index=False)

print(f"Duty list generated and saved to {output_file}")
