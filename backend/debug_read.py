import pandas as pd
import os

output_file_path = 'debug_output.txt'

try:
    script_dir = os.path.dirname(__file__)
    project_root = os.path.dirname(script_dir)
    file_path = os.path.join(project_root, 'sample.xls')

    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(f"Attempting to read: {file_path}\n")

        df = pd.read_excel(file_path, header=None)
        f.write("File read successfully!\n")

        # Replace NaN values with empty strings for cleaner output
        df = df.fillna('')

        f.write(str(df))

    print(f"Debug output written to {output_file_path}")

except Exception as e:
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(f"An error occurred: {e}\n")
    print(f"An error was written to {output_file_path}")
