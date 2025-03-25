import sys
import os
import csv
import json


class ExtractExcelODs:
    def __init__(self, input_folder, output_file):
        self.input_folder = input_folder
        self.output_file = output_file

        self.search_term = ["P", "l", "a", "t", "e", ":", "O", "F"]

    def remove_spaces_from_line(self, line):
        return line.replace(" ", "")
    
    def replace_tabs(self, line):
        return line.replace("\t", ",")

    def rewrite_file(self, file_path):
        try:
            with open(file_path, 'r') as file:
                lines = file.readlines()
                if len(lines) >= 52:
                    extracted_lines = lines[0:800]
                    line_number = 1
                    with open(self.output_file, 'w', newline='') as output:
                        writer = csv.writer(output)
                        row = []
                        for line in extracted_lines:
                            cleaned_line = self.remove_spaces_from_line(line).strip()
                            if cleaned_line:
                                if line_number % 2 != 0:
                                    row.append(f'{os.path.basename(file_path)[:-4]}={cleaned_line}')
                                else:
                                    row.append(cleaned_line)
                                    writer.writerow(row)
                                    row = []
                                line_number += 1
                else:
                    print(f"{os.path.basename(file_path)}: File does not contain enough lines (less than 52 lines).")
        except Exception as e:
            return
            # print(f"Error reading file {os.path.basename(file_path)}: {e}")

    def get_header(self, file_path, output_file):
        self.rewrite_file(file_path)
        try:
            with open(output_file, 'r') as file:
                lines = file.readlines()
                for i, line in enumerate(lines):
                    if all(char in line for char in self.search_term):
                        header_line = lines[i + 1].strip()
                        header_line_commas = self.replace_tabs(header_line)
                        print(f'{header_line_commas}')
                        return header_line  # Return the first header line found
        except Exception as e:
            print(f"Error reading file {os.path.basename(output_file)}: {e}")
        return None  # Return None if no header line is found

    def process_file(self, output_file):
        try:
            with open(output_file, 'r') as file:
                lines = file.readlines()
                for i, line in enumerate(lines):
                    if all(char in line for char in self.search_term):
                        # print(f'THIS IS : {line}')
                        for j in range(9):  # Process the next 9 lines
                            if i + j + 1 < len(lines) and j != 0:
                                # print(f'this is i: {i}, this is j: {j}')
                                next_line = lines[i + j + 1]
                                cleaned_line = self.remove_spaces_from_line(next_line).strip()
                                cleaner_line = self.replace_tabs(cleaned_line).strip()
                                if cleaner_line:
                                    print(cleaner_line)
                        break
        except Exception as e:
            print(f"Error reading file {os.path.basename(output_file)}: {e}")

    def process_files(self):
        try:
            if not os.path.isdir(self.input_folder):
                print("Error: Provided path is not a directory.")
                return
            # get header
            for filename in os.listdir(self.input_folder):
                file_path = os.path.join(self.input_folder, filename)
                if os.path.isfile(file_path):
                    self.get_header(file_path, self.output_file)
                break
            for filename in os.listdir(self.input_folder):
                file_path = os.path.join(self.input_folder, filename)
                if os.path.isfile(file_path):
                    self.rewrite_file(file_path)
                    self.process_file(self.output_file)
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    try:
        input_folder = sys.argv[1]
        user = os.getlogin()
        output_file = f'progresses/temp/{user}_temp.csv'
        processor = ExtractExcelODs(input_folder, output_file)
        processor.process_files()
    except IndexError:
        print("Error: Please provide a folder path as an argument.")
    except Exception as e:
        print(f"Error: {e}")