# Automate Table of Contents with Excel VBA
[![GitHub][github_badge]][github_link]
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Automate Table of Contents in Excel VBA is a program that generates a table of contents for an Excel workbook using VBA macros. This program eliminates the need for manually creating and updating the table of contents whenever changes are made to the workbook.

## Features

- Generates a table of contents for an Excel workbook.
- Automatically updates the table of contents when changes are made to the workbook.
- Provides clickable hyperlinks to navigate to individual sheets.

## Repository Structure

The repository is structured as follows:
```
├──AutoTOC.bas: The VBA code file that contains the macro for generating the table of contents.
├──AutoTOC Testing.xltm: A sample Excel workbook with the `TOC_Generator` macro on a "TOC" sheet to demonstrate how the program works.
├──LICENSE: The license file for the project.
├──README.md: provides an overview of this repository.
```

## Usage

To run the VBA program and generate a table of contents for your Excel workbook, please follow the steps below:

1. Download the VBA program from this [repository](https://github.com/MaxineXiong/AutomateTableOfContentsInExcel).

2. Open your Excel workbook.

3. Press `ALT+F11` to open the Visual Basic Editor.

4. In the Visual Basic Editor, go to `File > Import File` and select the downloaded **AutoTOC.bas** VBA code file.

5. Close the Visual Basic Editor.

6. Press `ALT+F8` to open the "Macro" dialog.

7. Select the `TOC_Generator` macro and click "Run".

8. Follow the program instructions step-by-step, and the program will generate a table of contents in the "TOC" sheet within the workbook.

9. To update the table of contents, simply run the `TOC_Generator` macro again.

Alternatively, you can test out the program by opening the **AutoTOC Testing.xltm** workbook and clicking the Table of Contents icon on the "TOC" sheet to see how it works.


![AutoTOC](https://github.com/MaxineXiong/AutomateTableOfContentsInExcel/assets/55864839/10a48f40-bd39-41a7-86d1-58ab9f9b53ba)


## Example

Let's assume we have an Excel workbook named *"MyWorkbook.xlsx"* with multiple sheets. To generate a table of contents for this workbook, follow the steps mentioned in the "Usage" section above. By following the program instructions step-by-step, the program will generate a table of contents in the "TOC" sheet within the workbook. To update the table of contents after making changes to the workbook, simply run the `TOC_Generator` macro again.

## License

This project is licensed under the [MIT License](https://choosealicense.com/licenses/mit/). Feel free to use, modify, and distribute the code in this repository.

[github_badge]: https://badgen.net/badge/icon/GitHub?icon=github&color=black&label
[github_link]: https://github.com/MaxineXiong
