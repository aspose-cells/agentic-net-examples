using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string inputPath = "input.xlsx";

        // Create LoadOptions and configure custom globalization settings
        LoadOptions loadOptions = new LoadOptions();
        // Example: use German culture (comma as decimal separator)
        loadOptions.CultureInfo = new CultureInfo("de-DE");

        // Load the workbook with the specified LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Demonstrate that the culture settings are applied
        Console.WriteLine("Cell A1 value with German culture: " + workbook.Worksheets[0].Cells["A1"].StringValue);

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}