import pandas as pd
import requests
import os


def fetch_data(url):
    response = requests.get(url)
    data = response.json()
    return data['users']

url = 'https://dummyjson.com/users/'
users = fetch_data(url)

print("First user's data to check the structure:")
print(users[0])


def process_data(users):
    df = pd.DataFrame(users)
    
    # Print columns to understand the data structure
    print("Available columns:", df.columns)
    
    # Initialize Latitude and Longitude columns
    df['Latitude'] = None
    df['Longitude'] = None
    
    # Loop through each user and extract Latitude and Longitude
    for i, user in df.iterrows():
        if 'address' in user and 'coordinates' in user['address']:
            df.at[i, 'Latitude'] = user['address']['coordinates']['lat']
            df.at[i, 'Longitude'] = user['address']['coordinates']['lng']
    
    # Convert 'birthDate' to 'DOB' in the desired format
    if 'birthDate' in df.columns:
        df['DOB'] = pd.to_datetime(df['birthDate']).dt.strftime('%d/%m/%Y')
    else:
        df['DOB'] = None
    
    # Rename columns to match the specified output format
    df = df.rename(columns={
        'firstName': 'Name',
        'email': 'Email Address',
        'age': 'Age',
        'gender': 'Gender',
        'phone': 'Phone'
    })
    
    # Define the columns needed for the output
    columns_needed = ['Name', 'Email Address', 'Age', 'Gender', 'Phone', 'Latitude', 'Longitude', 'DOB']
    
    # Handle missing columns
    for col in columns_needed:
        if col not in df.columns:
            df[col] = None
    
    # Select only the necessary columns
    df = df[columns_needed]
    
    return df

processed_data = process_data(users)
print("Processed data sample:")
print(processed_data.head())

def export_to_excel(df, file_name):
    try:
        # Export DataFrame to Excel
        df.to_excel(file_name, index=False)
        print(f"The Excel file has been successfully exported and saved as '{file_name}'.")
    except Exception as e:
        print(f"An error occurred while exporting to Excel: {e}")

print(f"Current working directory: {os.getcwd()}")

# Save the processed data to Excel in the current working directory
export_to_excel(processed_data, 'ProcessedUserData.xlsx')

