import pyexcel_ods, argparse

def parse_ods(file_path):
    data = pyexcel_ods.get_data(file_path)
    # 'data' is a dictionary where keys are sheet names and values are lists of rows
    return data

def custom_sort_key(value):
    if value == '':
        return 0  # Treat empty values as lowest priority
    else:
        return int(value)  # Convert numeric strings to integers for sorting

def generate_ordered_list(parsed_data, acc_column):
    # Extract data from the first sheet (assuming you're working with the first sheet)
    first_sheet_name = list(parsed_data.keys())[0]
    first_sheet_data = parsed_data[first_sheet_name]

    # Initialize indices for columns
    set_index = -1
    star_index = -1
    name_index = -1
    acc_01_index = -1
    gold_index = -1

    # Find index positions of SET, STAR, NAME, Acc_01, and GOLD columns
    if first_sheet_data:
        header_row = first_sheet_data[0]
        if 'SET' in header_row:
            set_index = header_row.index('SET')
        if 'STAR' in header_row:
            star_index = header_row.index('STAR')
        if 'NAME' in header_row:
            name_index = header_row.index('NAME')
        if 'Acc_01' in header_row:
            acc_01_index = header_row.index(acc_column)
        if 'GOLD' in header_row:
            gold_index = header_row.index('GOLD')

    # Filter rows where GOLD column does not contain 'X' and ensure all necessary columns exist
    filtered_data = []
    for row in first_sheet_data[1:]:
        if gold_index != -1 and gold_index < len(row) and row[gold_index] == 'X':
            continue  # Skip rows with 'X' in the GOLD column
        if set_index != -1 and star_index != -1 and name_index != -1 and acc_01_index != -1 and len(row) > max(set_index, star_index, name_index, acc_01_index):
            filtered_data.append(row)

    # Sort filtered data based on the STAR column (index star_index)
    if star_index != -1:
        sorted_data = sorted(filtered_data, key=lambda x: custom_sort_key(x[star_index]))
    else:
        sorted_data = filtered_data

    # Generate ordered list of SET, STAR, NAME, and Acc_01
    ordered_list = []
    for row in sorted_data:
        if set_index != -1 and star_index != -1 and name_index != -1 and acc_01_index != -1:
            ordered_list.append((row[set_index], row[star_index], row[name_index], row[acc_01_index]))

    return ordered_list

def format_output(ordered_list):
    formatted_output = []
    for item in ordered_list:
        if item[3] != '0' and item[3] != '':
            formatted_output.append(f"Set {item[0]} - {item[1]} Star - {item[2]} - x{item[3]}")
    return formatted_output

def parse_args():
    parser = argparse.ArgumentParser(description="Process ODS file and generate ordered list")
    parser.add_argument('acc_01_column', type=str, help="Name of the Acc_01 column")
    return parser.parse_args()

if __name__ == "__main__":

    # Parse command-line arguments
    args = parse_args()
    
    # Example usage
    file_path = 'MonGo_Inventory.ods'
    parsed_data = parse_ods(file_path)
    ordered_list = generate_ordered_list(parsed_data, args.acc_01_column)

    # Format ordered list output
    formatted_output = format_output(ordered_list)

    # Print formatted output
    for line in formatted_output:
        print(line)