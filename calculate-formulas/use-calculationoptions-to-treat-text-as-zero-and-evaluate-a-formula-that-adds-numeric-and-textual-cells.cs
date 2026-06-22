using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells: numeric value in A1 and textual value in A2
        sheet.Cells["A1"].PutValue(10);          // Numeric cell
        sheet.Cells["A2"].PutValue("sample");    // Text cell

        // Set calculation options.
        // IgnoreError = true prevents errors when non‑numeric text is used in arithmetic.
        // Aspose.Cells treats such text as zero during calculation when errors are ignored.
        CalculationOptions options = new CalculationOptions
        {
            IgnoreError = true
        };

        // Evaluate the formula that adds the two cells.
        // The text in A2 will be considered as zero.
        object result = sheet.CalculateFormula("=A1+A2", options);

        // Output the result.
        Console.WriteLine("Result of =A1+A2 (text treated as zero): " + result);

        // Save the workbook (optional demonstration of lifecycle rule usage).
        workbook.Save("Result.xlsx");
    }
}