using System;
using Aspose.Cells;

class LoadWorkbookExample
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Example option: skip formula parsing on open
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the workbook using the constructor that takes a file name and LoadOptions
        // Parameters are separated by commas as required
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet and read a cell value
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

        // Save the workbook to a new file
        workbook.Save("output.xlsx");
    }
}