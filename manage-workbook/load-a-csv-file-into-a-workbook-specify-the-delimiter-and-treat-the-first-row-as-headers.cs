using System;
using Aspose.Cells;

class LoadCsvWithHeaders
{
    static void Main()
    {
        // Path to the CSV file to be loaded
        string csvPath = "input.csv";

        // Create TxtLoadOptions for CSV format
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        // Specify the delimiter (e.g., comma)
        loadOptions.Separator = ',';
        // Treat the first row as header rows
        loadOptions.HeaderRowsCount = 1;
        // Optional: convert numeric strings to numeric cell types
        loadOptions.ConvertNumericData = true;

        // Load the CSV file into a workbook using the configured options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Access the first worksheet (optional demonstration)
        Worksheet sheet = workbook.Worksheets[0];
        // Example: display the value of the first data cell (A2) after the header
        Console.WriteLine("First data cell (A2): " + sheet.Cells["A2"].StringValue);

        // Save the workbook to an Excel file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}