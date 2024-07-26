import os
import shutil
import pandas as pd

def search_and_copy_files():
    # Check if the current file name is an AR file
    current_file_name = os.path.basename(__file__)
    if "AR Report" in current_file_name:
        new_folder_name = current_file_name.split("-")[0].strip()
    else:
        print("Please select the AR Report File, Then Run script")
        return

    # Prompt user to select the source folder
    source_folder_path = input("Please enter the source folder path: ")
    destination_folder_path = "C:\\Users\\sachin.mahadik\\Desktop\\Sub"

    # Create new folder with Customer Name in Submission folder
    destination_folder_path = os.path.join(destination_folder_path, new_folder_name)
    if not os.path.exists(destination_folder_path):
        os.makedirs(destination_folder_path)

    # Ensure the folder paths end with a backslash
    source_folder_path = os.path.join(source_folder_path, '')
    destination_folder_path = os.path.join(destination_folder_path, '')

    # Read the Excel file and find the last row
    excel_file = pd.ExcelFile(current_file_name)
    df = excel_file.parse(sheet_name='Sheet1')
    last_row = df.shape[0]

    # Loop through each cell in the specified range
    for index, row in df.iterrows():
        if index >= last_row:
            break
        file_name = row['F']
        file_found = False

        # Loop through each file in the source folder
        for file in os.listdir(source_folder_path):
            if file_name in file:
                # Construct the full source file path
                source_file = os.path.join(source_folder_path, file)

                # Copy the file to the destination folder
                shutil.copy2(source_file, destination_folder_path)
                file_found = True

                # Update the dataframe to indicate the file was copied
                df.at[index, 'P'] = 'Copied'
                break

        # If no file was found
        if not file_found:
            df.at[index, 'P'] = 'Not Found'

    # Save the updated dataframe back to the Excel file
    with pd.ExcelWriter(current_file_name, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Notify the user that the process is complete
    print("File search and copy process completed.")

if __name__ == "__main__":
    search_and_copy_files()
