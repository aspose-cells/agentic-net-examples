using System;
using Aspose.Cells;

class LoadCsvWithHeaders
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Create load options for a text file
        TxtLoadOptions loadOptions = new TxtLoadOptions();

        // Specify the delimiter (e.g., semicolon)
        loadOptions.Separator = ';';

        // Treat the first row as header rows
        loadOptions.HeaderRowsCount = 1;

        // Load the CSV into a workbook using the options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Example: read the first header cell
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Header A1: " + sheet.Cells["A1"].StringValue);

        // Save the workbook in Excel format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}