using System;
using Aspose.Cells;

class LoadWorkbookExample
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Create LoadOptions for XLSX format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Important: skip formula parsing on open if formulas are not needed immediately
        loadOptions.ParsingFormulaOnOpen = false;

        // Important: ignore overlapping useless shapes to reduce memory consumption
        loadOptions.IgnoreUselessShapes = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Access the first worksheet and read a sample cell
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].Value);

        // Example of modifying a workbook setting after load
        workbook.Settings.RepairLoad = true;

        // Save the workbook to a new file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}