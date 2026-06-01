using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Create LoadOptions and disable formula parsing on open
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false; // formulas will be loaded as text, not calculated

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Access the first worksheet and a cell that contains a formula
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];

        // Display the formula text (should be preserved) and the value (should be null/empty because not calculated)
        Console.WriteLine("Cell A1 Formula: " + cell.Formula);
        Console.WriteLine("Cell A1 Value after load (expected null or default): " + (cell.Value ?? "null"));

        // If you later need to calculate formulas manually, uncomment the following line:
        // workbook.CalculateFormula();

        // Save the workbook to verify that it remains unchanged
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}