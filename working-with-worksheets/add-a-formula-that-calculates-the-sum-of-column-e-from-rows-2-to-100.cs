// Title: Add a SUM(E2:E100) formula to cell F1 in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set SUM formula for range E2 to E100 | How to insert a column total formula with Aspose.Cells .NET | Assign Excel formula SUM(E2:E100) to a cell using Aspose.Cells API | Calculate workbook formulas immediately after adding SUM in Aspose.Cells C# | Save workbook after adding formula with Aspose.Cells for .NET
// Tags: Aspose.Cells add cell formula C# | Excel SUM range formula Aspose.Cells | calculate workbook formulas Aspose.Cells .NET | set formula result cell F1 Aspose | save workbook as xlsx Aspose.Cells

using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, places the formula "SUM(E2:E100)" into cell F1, forces immediate calculation of the formula, and saves the file as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a formula to calculate the sum of column E from rows 2 to 100
        // The result will be placed in cell F1 (you can change the target cell as required)
        Cell formulaCell = sheet.Cells["F1"];
        formulaCell.Formula = "SUM(E2:E100)";

        // Optionally calculate the formula immediately
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
