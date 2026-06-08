import pandas as pd
import os

# List all Excel files in uploads
files = [f for f in os.listdir('uploads') if f.endswith('.xlsx')]
print(f"Found {len(files)} Excel files in uploads folder:")
for f in files:
    print(f"  - {f}")

# Check the first processed file
if files:
    file_path = os.path.join('uploads', files[0])
    print(f"\nChecking file: {files[0]}")
    
    df = pd.read_excel(file_path)
    print(f"\nColumns: {list(df.columns)}")
    print(f"Shape: {df.shape}")
    
    # Check if it has SSH processing data
    has_ssh_data = all(col in df.columns for col in ['Power Status', 'Action', 'Power Levels'])
    print(f"\nHas SSH processing data: {has_ssh_data}")
    
    if has_ssh_data:
        print("\nSample data:")
        print(df[['ID', 'Power Status', 'Action', 'Power Levels']].head(5).to_string())
        
        # Count by action
        print("\n\nAction counts:")
        print(df['Action'].value_counts())
        
        print("\n\nPower Status counts:")
        print(df['Power Status'].value_counts())

