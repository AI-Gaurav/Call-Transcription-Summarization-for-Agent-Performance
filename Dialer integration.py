import os
import pandas as pd
import requests
from urllib.parse import urlparse
import glob

# Configuration
download_root = "D:\\Dialer Audio Download"
audio_folder = os.path.join(download_root, "audio_files")
output_excel = os.path.join(download_root, "audio_files_details.xlsx")

# Create directories if they don't exist
os.makedirs(audio_folder, exist_ok=True)

# Find CSV file in the directory
csv_files = glob.glob(os.path.join(download_root, "*.csv"))

if not csv_files:
    raise FileNotFoundError(f"No CSV files found in {download_root}")

# Use the first CSV file found
csv_path = csv_files[0]
print(f"Found CSV file: {csv_path}")

try:
    # Read the CSV file with different encoding options
    try:
        df = pd.read_csv(csv_path)
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(csv_path, encoding='latin1')
        except Exception as e:
            df = pd.read_csv(csv_path, encoding='utf-16')

    print("Available columns in CSV:", df.columns.tolist())

    # Try to find audio link column (case insensitive)
    audio_column = None
    possible_names = ['R', 'Recording', 'Audio', 'URL', 'Link']  # Common column names
    
    for col in df.columns:
        if any(name.lower() == col.lower() for name in possible_names):
            audio_column = col
            break
    
    if not audio_column:
        # If still not found, try to find any column containing URLs
        for col in df.columns:
            if df[col].astype(str).str.contains('http').any():
                audio_column = col
                break
    
    if not audio_column:
        raise ValueError(f"No audio link column found. Please check your CSV file. Available columns: {df.columns.tolist()}")
    
    print(f"Using column '{audio_column}' for audio links")

    # Get unique audio links and remove NaN values
    audio_links = df[audio_column].dropna().unique()
    print(f"Found {len(audio_links)} unique audio links to download")

    # Prepare data for Excel output
    output_data = []
    
    for index, link in enumerate(audio_links, start=1):
        try:
            # Skip if empty link
            if pd.isna(link) or str(link).strip() == "":
                continue
                
            link = str(link).strip()  # Clean the link
            
            # Extract filename from URL
            parsed_url = urlparse(link)
            filename = os.path.basename(parsed_url.path)
            
            # If filename is empty or doesn't have an extension, create one
            if not filename or '.' not in filename:
                filename = f"audio_{index}.mp3"
            else:
                # Clean filename by removing query parameters if any
                filename = filename.split('?')[0]
            
            file_path = os.path.join(audio_folder, filename)
            
            # Download the file with timeout and headers
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(link, stream=True, timeout=30, headers=headers)
            response.raise_for_status()
            
            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            # Add to output data
            output_data.append({
                'Original Link': link,
                'File Name': filename,
                'File Path': file_path,
                'Status': 'Downloaded'
            })
            
            print(f"Successfully downloaded: {filename}")
            
        except Exception as e:
            output_data.append({
                'Original Link': link,
                'File Name': '',
                'File Path': '',
                'Status': f'Error: {str(e)}'
            })
            print(f"Error downloading {link}: {e}")
    
    # Create output DataFrame
    output_df = pd.DataFrame(output_data)
    
    # Save to Excel
    output_df.to_excel(output_excel, index=False)
    print(f"\nAll downloads completed. Details saved to: {output_excel}")

except Exception as e:
    print(f"An error occurred: {str(e)}")
    input("Press Enter to exit...")  # Keep window open to see error
    raise