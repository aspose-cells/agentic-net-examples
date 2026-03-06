using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Numbers;

class Program
{
    static void Main()
    {
        // Path to the Apple Numbers file to be opened
        string numbersFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.numbers");

        Workbook workbook;

        if (File.Exists(numbersFilePath))
        {
            // Create load options for Numbers files
            NumbersLoadOptions loadOptions = new NumbersLoadOptions
            {
                // Example setting: load each sheet as a separate table
                LoadTableType = LoadNumbersTableType.OneTablePerSheet
            };

            // Load the Numbers spreadsheet into a Workbook instance
            workbook = new Workbook(numbersFilePath, loadOptions);
        }
        else
        {
            // If the .numbers file is not found, create a simple workbook instead
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";
            sheet.Cells["A1"].PutValue("Sample data - .numbers file not found.");
        }

        // Iterate through worksheets and output their names (demonstration purpose)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");
        }

        // Save the loaded workbook to an Excel file (XLSX format)
        string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertedFromNumbers.xlsx");
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to: {outputPath}");
    }
}