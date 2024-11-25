import pandas as pd

# Load the Excel file
df = pd.read_excel(r'C:\Users\91975\Desktop\Project_excel\generated_data.xlsx')

# Function to classify rows (this should be defined based on previous code)
def classify_row(row):
    time_taken = row['Time Taken']
    priority = row['priority']
    row_type = row['Type']

    if priority == 'p1c' and row_type == 'type1':
        if time_taken < 3:
            return '<3 hours'
        elif time_taken < 8:
            return '<8 hours'
        else:
            return '>8 hours'
    if priority == 'p1s' and row_type == 'type1':
            if time_taken < 4:
                return '<3 hours'
            elif time_taken < 8:
                return '<8 hours'
            else:
                return '>8 hours'
        
    if priority == 'p2' and row_type == 'type1':
            if time_taken < 24:
                return '1 day'
            elif time_taken < 48:
                return '<2 days'
            else:
                return '>2 days'
    if priority == 'p3' and row_type == 'type1':
            if time_taken < 48:
                return '2 days'
            elif time_taken < 96:
                return '<4 days'
    else:
            return '>4 days'
        
    
    
    if priority == 'p1c' and row_type == 'type2':
        if time_taken < 3:
            return '<3 hours'
        elif time_taken < 8:
            return '<8 hours'
        else:
            return '>8 hours'
    if priority == 'p1s' and row_type == 'type2':
            if time_taken < 4:
                return '<3 hours'
            elif time_taken < 8:
                return '<8 hou rs'
    else:
            return '>8 hours'
        
    if priority == 'p2' and row_type == 'type2':
            if time_taken < 24:
                return '1 day'
            elif time_taken < 48:
                return '<2 days'
    else:
            return '>2 days'
    if priority == 'p3' and row_type == 'type2':
            if time_taken < 48:
                return '2 days'
            elif time_taken < 96:
                return '<4 days'
    else:
            return '>4 days'
    # Add classification rules for other priority/type combinations
    # Similar for p1s, p2, p3 based on the previous conditions
    return 'None'  # Default case for missing classification

# Apply the classification function to each row
df['Classification'] = df.apply(classify_row, axis=1)

# Step 2: Calculate the percentages for each priority and type
def calculate_percentages(df):
    summary = []
    # Group by 'priority', 'Type', 'Classification'
    grouped = df.groupby(['priority', 'Type', 'Classification']).size().reset_index(name='Count')
    
    # Calculate total for each priority and type combination
    total_grouped = df.groupby(['priority', 'Type']).size().reset_index(name='Total')

    # Merge to get the total count
    merged = pd.merge(grouped, total_grouped, on=['priority', 'Type'])

    # Calculate percentage
    merged['Percentage'] = (merged['Count'] / merged['Total']) * 100
    return merged

# Step 3: Calculate percentages per month
def calculate_percentages_per_month(df):
    summary = []
    # Group by 'priority', 'Type', 'Classification', 'Month'
    grouped = df.groupby(['priority', 'Type', 'Classification', 'Month']).size().reset_index(name='Count')
    
    # Calculate total for each priority and type combination for each month
    total_grouped = df.groupby(['priority', 'Type', 'Month']).size().reset_index(name='Total')

    # Merge to get the total count
    merged = pd.merge(grouped, total_grouped, on=['priority', 'Type', 'Month'])

    # Calculate percentage
    merged['Percentage'] = (merged['Count'] / merged['Total']) * 100
    return merged

# Step 4: Repeat the same for each group
def calculate_percentages_per_group(df):
    summary = []
    # Group by 'priority', 'Type', 'Classification', 'Group'
    grouped = df.groupby(['priority', 'Type', 'Classification', 'Group']).size().reset_index(name='Count')
    
    # Calculate total for each priority and type combination for each group
    total_grouped = df.groupby(['priority', 'Type', 'Group']).size().reset_index(name='Total')

    # Merge to get the total count
    merged = pd.merge(grouped, total_grouped, on=['priority', 'Type', 'Group'])

    # Calculate percentage
    merged['Percentage'] = (merged['Count'] / merged['Total']) * 100
    return merged

# Step 5: Repeat the same for "reported" records
def calculate_percentages_for_reported(df):
    # Filter records where 'Report in' is 1
    reported_df = df[df['Report in'] == 1]

    return calculate_percentages(reported_df)

# Step 6: Calculate percentages for each group and month for reported records
def calculate_percentages_per_group_and_month_for_reported(df):
    reported_df = df[df['Report in'] == 1]

    # Per group
    group_percentages = calculate_percentages_per_group(reported_df)
    
    # Per month
    month_percentages = calculate_percentages_per_month(reported_df)
    
    return group_percentages, month_percentages

# Generate results
total_percentages = calculate_percentages(df)
month_percentages = calculate_percentages_per_month(df)
group_percentages = calculate_percentages_per_group(df)
reported_percentages = calculate_percentages_for_reported(df)
reported_group_percentages, reported_month_percentages = calculate_percentages_per_group_and_month_for_reported(df)

# Step 7: Save all the results to a new Excel sheet
with pd.ExcelWriter('updated_excel_file_with_percentages.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Original with Classification', index=False)
    total_percentages.to_excel(writer, sheet_name='Total Percentages', index=False)
    month_percentages.to_excel(writer, sheet_name='Monthly Percentages', index=False)
    group_percentages.to_excel(writer, sheet_name='Group Percentages', index=False)
    reported_percentages.to_excel(writer, sheet_name='Reported Percentages', index=False)
    reported_group_percentages.to_excel(writer, sheet_name='Reported Group Percentages', index=False)
    reported_month_percentages.to_excel(writer, sheet_name='Reported Monthly Percentages', index=False)

print("Updated data with classifications and percentages saved to updated_excel_file_with_percentages.xlsx")
