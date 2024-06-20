# Load existing data if the file exists
try:
        df = pd.read_excel('receivers.xlsx')
except FileNotFoundError:
        df = pd.DataFrame()

    # Append the new data to the DataFrame
new_data = pd.DataFrame({'Name': [recipient_name], 'Email': [recipient_email], 'Subject': [subject]})
df = pd.concat([df, new_data], ignore_index=True)

    # Save the DataFrame to Excel
df.to_excel('receivers.xlsx', index=False, engine='openpyxl')