using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create load options for XLSX files
        LoadOptions loadOptions = new LoadOptions();
        // Ensure that formulas are parsed when the workbook is opened
        loadOptions.ParsingFormulaOnOpen = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Assume cell A1 contains a formula.
        // Use the FORMULATEXT function to retrieve the formula text of A1 and place it in B1.
        cells["B1"].Formula = "=FORMULATEXT(A1)";

        // Calculate all formulas so that the FORMULATEXT result is evaluated
        workbook.CalculateFormula();

        // Display the original formula in A1 and the text returned by FORMULATEXT in B1
        Console.WriteLine("Cell A1 formula: " + cells["A1"].Formula);
        Console.WriteLine("Cell B1 (FORMULATEXT) result: " + cells["B1"].StringValue);

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}