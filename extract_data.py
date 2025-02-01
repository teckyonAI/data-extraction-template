import pandas as pd
from bs4 import BeautifulSoup

def extract_data_from_html(file_path):
    # Read the HTML file
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    # Parse the HTML with BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Locate the table
    table = soup.find('table', {'id': 'data-table'})
    if not table:
        raise Exception("Data table not found in the HTML.")
    
    # Extract header names
    headers = [th.get_text().strip() for th in table.find_all('th')]
    
    # Extract rows
    data = []
    for row in table.find_all('tr')[1:]:  # Skip the header row
        cols = row.find_all('td')
        if cols:
            row_data = [col.get_text().strip() for col in cols]
            data.append(row_data)
    
    return headers, data

def save_data_to_excel(headers, data, output_file):
    # Create a DataFrame and save it to Excel
    df = pd.DataFrame(data, columns=headers)
    df.to_excel(output_file, index=False)
    print(f"Data successfully saved to {output_file}")

def main():
    html_file = 'program_output.html'
    output_file = 'output.xlsx'
    
    try:
        headers, data = extract_data_from_html(html_file)
        save_data_to_excel(headers, data, output_file)
    except Exception as e:
        print(f"Error during extraction: {e}")

if __name__ == "__main__":
    main()
