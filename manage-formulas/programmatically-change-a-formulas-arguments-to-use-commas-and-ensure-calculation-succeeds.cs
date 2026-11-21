// Title: Replace semicolon separators with commas in an Excel formula and recalculate the workbook using Aspose.Cells for .NET
// Common Searches: how to change formula argument separator from semicolon to comma with Aspose.Cells C# | Aspose.Cells replace semicolon delimiters in Excel formula programmatically | force recalculation after modifying a cell formula using Aspose.Cells .NET | C# code to correct Excel formula delimiters and recalculate workbook with Aspose.Cells | update Excel cell formula separators and recalc workbook Aspose.Cells example
// Tags: Aspose.Cells formula delimiter replacement | C# Excel formula argument conversion | Aspose.Cells workbook recalculation | programmatic cell formula edit .NET | Excel semicolon to comma formula fix

using Aspose.Cells;
using System;

// The example loads an Excel workbook, checks cell A1 for a formula, replaces any semicolon argument separators with commas, reassigns the corrected formula, forces full workbook recalculation, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Example: target cell that contains a formula
        Cell targetCell = sheet.Cells["A1"];

        // Check if the cell actually contains a formula
        if (targetCell.IsFormula)
        {
            // Retrieve the existing formula text
            string originalFormula = targetCell.Formula;

            // Replace argument separators (e.g., semicolons) with commas
            string correctedFormula = originalFormula.Replace(";", ",");

            // Assign the corrected formula back to the cell
            targetCell.Formula = correctedFormula;
        }

        // Force recalculation of all formulas to ensure the change takes effect
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
