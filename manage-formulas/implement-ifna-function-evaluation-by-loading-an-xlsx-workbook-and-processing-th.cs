using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing workbook (replace with your actual file)
        string inputPath = "input.xlsx";

        // Load the workbook with formula parsing enabled
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set up values that will cause an error in the division
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(0); // Division by zero will produce #DIV/0!

        // Apply IFNA formula: if the division results in an error, return "DivZero"
        cells["C1"].Formula = "=IFNA(A1/B1, \"DivZero\")";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the evaluated result of the IFNA function
        Console.WriteLine("Result of IFNA in C1: " + cells["C1"].StringValue);

        // Save the workbook with the evaluated result (optional)
        workbook.Save("output.xlsx");
    }
}