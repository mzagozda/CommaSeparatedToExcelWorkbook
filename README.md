# Description
The application converts CSV file to Excel file. The application expects two arguments: the path to the input CSV file and the path for the output Excel file. Defaults are provided if no arguments are supplied.
Reading CSV: The CsvHelper library is used to read the CSV file. It supports reading into a dynamic object for flexibility.
Writing to Excel: The ClosedXML library creates an Excel workbook, adds data to cells, and saves the file.
# Build and Run
Build your project in Visual Studio and run it. You can also run it from the command line or terminal by navigating to the output directory and running:

```bash
dotnet run "path/to/input.csv" "path/to/output.xlsx"
This approach provides a basic structure. You might need to adapt the code depending on the specifics of your CSV file (like handling different data types, handling more complex CSV structures, etc.).
```
