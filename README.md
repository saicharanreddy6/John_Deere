# XML Processor

This repository contains a Python application that processes XML files by extracting text data from sections and subsections. The project includes functionality for converting table elements to Markdown and outputting the results in a CSV file.

## Project Structure

- **src/**: Contains source code.
  - `xml_processor.py`: Main script.
  - `utils.py`: Helper functions.
- **tests/**: Unit tests.
  - `test_xml_processor.py`: Test cases.
- **data/**: Input and output data files.
  - `omdxe11330.xml`: Sample XML input.
  - `output.csv`: Output file.
- **docs/**: Documentation.
- **.gitignore**: Excludes unnecessary files.
- **requirements.txt**: Dependencies.
- **setup.py**: Packaging configuration.


## Features

- **XML Parsing:** Uses Python’s built-in `xml.etree.ElementTree` module to parse XML files.
- **Data Extraction:** Extracts text from `<omsection>` elements and their nested `<block>` elements.
- **Markdown Conversion:** Converts XML `<table>` elements into Markdown formatted tables.
- **CSV Output:** Saves the extracted data into a CSV file.
- **Unit Testing:** Includes unit tests (using Python’s `unittest` framework) to ensure extraction and CSV generation work as expected.

## File Structure

- **xml_processor.py:** Contains the `XMLProcessor` class along with the main execution logic that processes the XML file and generates the CSV.
- **test_xml_processor.py:** Contains unit tests to verify the functionality of the XMLProcessor class.

## Prerequisites

- Python 3.x
- Import pandas
- No additional libraries are required (only standard Python libraries are used).

## How to Use

1. **Prepare Your XML File:**
   - Place your XML file (for example, `omdxe11330.xml`) in the same directory as `xml_processor.py`.

2. **Run the Main Application:**
   - Execute the following command to process the XML file and generate the CSV:
     ```bash
     python xml_processor.py
     ```
   - The output CSV file (`output.csv`) will be created in the same directory.

3. **Run Unit Tests:**
   - To run the unit tests, execute:
     ```bash
     python test_xml_processor.py
     ```
   - All tests should pass, confirming that the extraction and CSV storage functionalities are working correctly.

## Logging

The application uses Python's `logging` module to provide informative messages about the XML parsing process and CSV file generation. These logs are output to the console.


## Usage

1. Place your XML file in the `data/` directory.
2. Run the main script from the `src/` directory:
