# **Content produced for Data Analytics and Jounral Entry Testing**
Analytics Tools is a VBA library for dealing with the formatting of several commonly utilized Accounting Tools

## Installation

1. Download the VBA files from the GitHub repository:

```bash
https://github.com/matt-chinchilla/AnalyticsTools/tree/main
```
2. Open Excel and press `Alt + F11` to open the VBA editor.
3. Import the downloaded `.bas` files into your Excel workbook.
4. If import of `.bas` is not an option, try pasting as a `.txt` file instead

## Usage

After installing the scripts in your Excel workbook, you can run the `SageMacro` subroutine to process your ledger. Here is an example of how to call `SageMacro`:

```vba
Sub RunSageMacro()
 SageMacro
End Sub
```
### **The macro specified performs the following steps:**

* Turns off screen updating and event handling for performance optimization.
* Displays gridlines for better visibility.
* Searches for specific header keywords in the ledger and deletes all rows above the first instance of these headers.
* Formats the worksheet, including renaming headers and inserting new columns.
* Filters out unwanted data and fills in data gaps.
* Creates a summary sheet with account totals.
* Modifies the summary sheet's format for clarity.
* Re-enables screen updating and event handling.

### The _SageMacro_ calls several helper subroutines such as:
* _DeleteRowsAboveKeywords_ 
* _FormatWorksheet_ 
* _FilterAndFillData_
* _MakeFinal_
* _MakeSummarySheet_ and
* _GetAccountInfo_

Each of these has a specific role in the processing of the ledger.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please ensure you provide detailed comments with you code and explain the rationale behind your changes.
