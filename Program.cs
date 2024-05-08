using System;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using ClosedXML.Excel;
// ReSharper disable SuggestVarOrType_BuiltInTypes

string csvFilePath = args.Length > 0 ? args[0] : "input.csv";
string excelFilePath = args.Length > 1 ? args[1] : "output.xlsx";

try
{
    // Read CSV file using CsvHelper
    using (var reader = new StreamReader(csvFilePath))
    using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = "," }))
    {
        // Get all records from the CSV file
        var records = csv.GetRecords<dynamic>();

        // Create a new Excel workbook
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Sheet 1");

            // Dynamically create columns based on the CSV header
            bool headerProcessed = false;
            int row = 1;

            foreach (var record in records)
            {
                int col = 1;
                foreach (var property in (IDictionary<string, object>)record)
                {
                    if (!headerProcessed)
                    {
                        worksheet.Cell(1, col).Value = property.Key;
                    }
                    worksheet.Cell(row + 1, col).Value = (XLCellValue)property.Value;
                    col++;
                }
                if (!headerProcessed)
                {
                    headerProcessed = true;
                }
                row++;
            }

            // Save the workbook to a file
            workbook.SaveAs(excelFilePath);
        }
    }

    Console.WriteLine("Excel file created successfully!");
}
catch (Exception ex)
{
    Console.WriteLine("Error: " + ex.Message);
}