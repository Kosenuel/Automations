# This SearchForAndOutput2.0.py script was created by **Emmanuel Okose**. #

* This file receives two inputs

## Description

This script receives two excel files by asking the user to input the file paths of the files at intervals, it also requests that the user specifies the columns to search for matching data, and asks the for the name to give the output file (it's going to generate). 

This script then searches for matches between the columns specified by the user in the two files provided {File one = Searcher, File two = file to be searched :) } and creates a new Excel file containing the results.

## Usage
- Clone the repository or download this script `SearchForAndOutput.py`.
- Ensure that you have installed the Python SDK on your system, added the path to environmental variables and have the `openpyxl` library installed. If not, install python from this link [Click here](https://www.python.org/ftp/python/3.12.2/python-3.12.2-amd64.exe). Install properly with admin privileges and add to path (if necessary), then install the openpyxl library using pip: `pip install openpyxl`

- Run the script in a Python environment.
- Follow the prompts from the script to input the file paths { for the searcher and searchee files :) }, column numbers, and give the output file a name as prompted.
- Press Enter, and the script will now search for matching data between the specified columns in the two Excel files and create an output Excel file with the name you specified containing the result.

* ***REMEMBER: THIS SCRIPT CAN'T PROPERLY READ DATA FROM A COLUMN/CELLS THAT HAS IT'S VALUE AS A RESULT OF A FORMULAR***.

* ***THIS SCRIPT CANNOT ALSO READ AN EXCEL FILE THAT IS TAGGED 'INTERNAL' OR 'CONFIDENTIAL', IT MUST BE 'PUBLIC'***

## Requirements

- Python 3.x
- openpyxl library

## Example

Here's an example of how to use the script:
1. Enter the path to the first Excel file: path/to/first file.xlsx
2. Enter the path to the second Excel file: path/to/the_second_file.xlsx
3. Enter the column number in the first file to use for searching (1-based): 2
Enter the column to be searched in the second file (1-based): 1
Enter the column to get content from the second file: 3
Enter the desired name for the output file: This_is_my_new_file.xlsx


This script will then search for matching data of cells in column 2 of the first file inside column 1 of the second file. It will then output the results to a new Excel file named "This_is_my_new_file.xlsx".

## Contributing

Contributions are welcome! IF you find any issues or have any suggestions for improvement, feel free to open an issue, submit a pull request or reach out to me @ immanuel.okose@gmail.com

## License

This project is licensed under the OptimusBank License [just kidding. I work here]