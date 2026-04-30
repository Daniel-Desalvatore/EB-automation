import re

def filter_log_file(input_filepath, output_filepath):
    """
    Reads a log file, filters lines containing "DEBUG - Action for transaction ID",
    and writes the filtered lines to a new file.

    Args:
        input_filepath (str): The path to the input log file.
        output_filepath (str): The path to the output file where filtered lines will be saved.
    """
    try:
        with open(input_filepath, 'r') as infile, open(output_filepath, 'w') as outfile:
            for line in infile:
                if "DEBUG - Action for transaction ID" in line:
                    outfile.write(line)
        print(f"Filtered log saved to: {output_filepath}")
    except FileNotFoundError:
        print(f"Error: The file '{input_filepath}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    input_log_file = "eb logs.txt"  # Replace with your input log file name
    output_log_file = "filtered_log.txt" # Replace with your desired output file name

  

    filter_log_file(input_log_file, output_log_file)