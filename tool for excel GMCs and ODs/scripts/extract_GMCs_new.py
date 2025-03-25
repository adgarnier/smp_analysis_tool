import sys
import os
import csv
import json

user = os.getlogin()

class ExtractExcelODs:
    def __init__(self, input_folder, output_file):
        self.input_folder = input_folder
        self.output_file = output_file

        self.user = os.getlogin()
        with open(f'configs/config{self.user}.json', 'r') as file:
            config = json.load(file)

        self.select_test = config['select_test']
        # print(f'{self.select_test}')

        if self.select_test == 'Default':
            self.search_term = [":", "G", "r", "o", "u", "p", "S", "a", "m", "p", "l", "e", "s"]
            self.extract_range_start = 1
            self.extract_range_amount = 58
            self.extract_range_step = 8
        elif self.select_test == 'VzV':
            self.search_term = [":", "G", "r", "o", "u", "p", "P", "s", ".", "C", "t", "l", "s", "-"]
            self.extract_range_start = 25
            self.extract_range_amount = 66
            self.extract_range_step = 8
        elif self.select_test == 'GSK ELLA':
            self.search_term = [":", "G", "r", "o", "u", "p", "S", "a", "m", "p", "l", "e", "s"]
            self.extract_range_start = 1
            self.extract_range_amount = 72
            self.extract_range_step = 10

    def remove_spaces_from_line_na(self, line):
        if "p" in line:
            return line
        else:
            return line.replace(" ", "NA")
    
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
                            cleaned_line = self.remove_spaces_from_line_na(line).strip()
                            if cleaned_line:
                                if line_number % 2 != 0:
                                    row.append(f'{os.path.basename(file_path)[:-4]}=\t{cleaned_line}')
                                else:
                                    row.append(f'{cleaned_line}')
                                    writer.writerow(row)
                                    row = []
                                line_number += 1
                else:
                    print(f"{os.path.basename(file_path)}: File does not contain enough lines (less than 52 lines).")
        except Exception as e:
            # return
            print(f"Error reading file {os.path.basename(file_path)}: {e}")

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
            # print(self.header_line)
            with open(output_file, 'r') as file:
                lines = file.readlines()
                for i, line in enumerate(lines):
                    if all(char in line for char in self.search_term):
                        # header_line = lines[i +1].strip()
                        # print(f'THIS IS : {line} HEADER: {header_line}')
                        # print(header_line)
                        for j in range(self.extract_range_start, self.extract_range_amount, self.extract_range_step):  # Process the next 9 lines
                            if i + j + 1 < len(lines) and j != 0:
                                # print(f'this is i: {i}, this is j: {j}')
                                next_line = lines[i + j + 1]
                                cleaned_line = (next_line).strip()
                                cleaned_line_commas = self.replace_tabs(cleaned_line)
                                if cleaned_line_commas:
                                    print(cleaned_line_commas)
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
            # all data
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
        output_file = f'progresses/temp/{user}_temp.csv'
        processor = ExtractExcelODs(input_folder, output_file)
        processor.process_files()
    except IndexError:
        print("Error: Please provide a folder path as an argument.")
    except Exception as e:
        print(f"Error: {e}")