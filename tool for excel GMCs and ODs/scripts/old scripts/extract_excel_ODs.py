import sys
import os

class ExtractExcelODs:
    def __init__(self, input_folder):
        self.input_folder = input_folder

    def remove_spaces_from_line(self, line):
        return line.replace(" ", "")

    def process_file(self, file_path):
        try:
            with open(file_path, 'r') as file:
                lines = file.readlines()
                if len(lines) >= 52:
                    extracted_lines = lines[36:52]
                    line_number = 1
                    for line in extracted_lines:
                        cleaned_line = self.remove_spaces_from_line(line).strip()
                        ### problem with cleaned_line
                        if cleaned_line:
                            if line_number % 2 != 0:
                                print(f'{os.path.basename(file_path)}: {cleaned_line}', end='\t')
                            else:
                                print(cleaned_line)
                            line_number += 1
                else:
                    print(f"{os.path.basename(file_path)}: File does not contain enough lines (less than 52 lines).")
        except Exception as e:
            print(f"Error reading file {os.path.basename(file_path)}: {e}")

    def process_files(self):
        try:
            if not os.path.isdir(self.input_folder):
                print("Error: Provided path is not a directory.")
                return

            header = "NA\tNA\t1\t2\t3\t4\t5\t6\t7\t8\t9\t10\t11\t12"
            print(header)

            for filename in os.listdir(self.input_folder):
                file_path = os.path.join(self.input_folder, filename)
                if os.path.isfile(file_path):
                    self.process_file(file_path)
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    try:
        input_folder = sys.argv[1]
        processor = ExtractExcelODs(input_folder)
        processor.process_files()
    except IndexError:
        print("Error: Please provide a folder path as an argument.")
    except Exception as e:
        print(f"Error: {e}")