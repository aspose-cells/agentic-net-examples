using System;
using Aspose.Cells;

class SetComplexFormulaForNamedRange
{
    static void Main()
    {
        // Paths for input and output XLSX files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (A1:A10 and B1:B10)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i + 1);          // Column A: 1,2,...,10
            cells[i, 1].PutValue((i + 1) * 2);    // Column B: 2,4,...,20
        }

        // Add a named formula (named range that refers to a formula)
        int nameIdx = workbook.Worksheets.Names.Add("MyComplexFormula");
        Name namedFormula = workbook.Worksheets.Names[nameIdx];

        // Complex formula: SUM of A1:A10 multiplied by AVERAGE of B1:B10
        string complexFormula = "=SUM(Sheet1!$A$1:$A$10)*AVERAGE(Sheet1!$B$1:$B$10)";

        // Set the reference of the name to the formula (using SetRefersTo)
        // Parameters: formula string, isR1C1 = false, isLocal = false
        namedFormula.SetRefersTo(complexFormula, false, false);

        // Use the named formula in a cell to verify it works
        cells["D1"].Formula = "=MyComplexFormula";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save(outputPath);
    }
}