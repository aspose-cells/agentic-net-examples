using System;
using Aspose.Cells;

class EnsureFormulaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Set formulas using a helper that guarantees a leading '='
            SetCellFormula(cells["B1"], "SUM(A1,A2)");          // missing '='
            SetCellFormula(cells["B2"], "=AVERAGE(A1:A2)");    // already has '='

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated results
            Console.WriteLine("B1 value: " + cells["B1"].Value);
            Console.WriteLine("B2 value: " + cells["B2"].Value);

            // Save the workbook to a file
            string outputPath = "EnsuredFormulas.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method that ensures the formula string starts with '=' before setting it
    static void SetCellFormula(Cell cell, string formula)
    {
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        // Set the formula directly; Aspose.Cells will compute the value later
        cell.Formula = formula;
    }
}