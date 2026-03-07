using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Condition that determines whether formulas should be parsed when the workbook is opened
        bool parseFormulas = false; // Change to true to enable formula parsing

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Apply the condition to set the ParsingFormulaOnOpen property
        if (parseFormulas)
        {
            loadOptions.ParsingFormulaOnOpen = true;
        }
        else
        {
            loadOptions.ParsingFormulaOnOpen = false;
        }

        // Load the workbook using the specified load options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Access the first worksheet and display information about cell A1
        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];
        Console.WriteLine("Cell A1 Formula: " + cell.Formula);
        Console.WriteLine("Cell A1 Value: " + cell.Value);

        // Save the workbook (no modifications made, just demonstrating load/save)
        workbook.Save("output.xlsx");
    }
}