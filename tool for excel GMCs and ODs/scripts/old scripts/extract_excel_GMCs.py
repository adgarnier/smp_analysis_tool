import sys
import os

def remove_spaces_from_line(line):
    if "p" in line:
        return line
    else:
        return line.replace(" ", "NA")

def print_help():
    help_message = """
    Usage: python extract_excel_GMCs.py <input_folder> [start]

    This script processes excel files in the specified input folder. It extracts GMC lines
    for each sample and prints the cleaned lines to the console. Paste output in excel and deliminate.

    Arguments:
    <input_folder>  The path to the folder containing the text files to be processed.
    [start]         (Optional) The line number to start processing from. Default is 224.

    Options:
    -h              Show this help message and exit.
    """
    print(help_message)

def main():
    if len(sys.argv) > 1 and sys.argv[1] == '-h':
        print_help()
        sys.exit(0)

    try:
        input_folder = sys.argv[1]
        if not os.path.isdir(input_folder):
            print("Error: Provided path is not a directory.")
            sys.exit(1)

        # Check for optional start argument
        if len(sys.argv) > 2:
            try:
                start = int(sys.argv[2])
            except ValueError:
                print("Error: Start value must be an integer.")
                sys.exit(1)
        else:
            start = 224  # Default start value

        for filename in os.listdir(input_folder):
            file_path = os.path.join(input_folder, filename)
            if os.path.isfile(file_path):
                try:
                    with open(file_path, 'r') as file:
                        lines = file.readlines()
                        if len(lines) >= 258:  # Ensure there are enough lines in the file
                            step = 16
                            target_lines_indices = [start - 2] + [start + step * i for i in range(8)]
                            extracted_lines = [lines[i] for i in target_lines_indices]

                            for i, line in enumerate(extracted_lines):
                                cleaned_line = remove_spaces_from_line(line).rstrip()
                                if cleaned_line:  # Check if the cleaned line is not empty
                                    print(cleaned_line)

                            print("\n")
                        else:
                            print(f"{filename}: File does not contain enough lines (less than 258 lines).")
                except Exception as e:
                    print(f"Error reading file {filename}: {e}")

    except IndexError:
        print("Error: Please provide a folder path as an argument.")
        print_help()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()