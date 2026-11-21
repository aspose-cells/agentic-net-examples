// Title: Merge cells A1:D1 and apply a SUM formula using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge a range and set a SUM formula | C# example for merging cells A1 to D1 and adding a total calculation in Aspose.Cells | How to assign a formula to a merged cell with Aspose.Cells .NET | Create a workbook, merge cells, and set a formula using Aspose.Cells in C#
// Tags: merge cells range Aspose.Cells C# | set SUM formula merged cell Aspose.Cells | Aspose.Cells workbook creation and cell merging | apply formula to merged range .NET | Aspose.Cells cell merging with formula assignment

using Aspose.Cells;

// The sample creates a new workbook, merges cells A1:D1 on the first worksheet, assigns the formula "=SUM(E1:E10)" to the merged cell, and saves the file as MergedCellTotal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells from column A to D in the first row (A1:D1)
        // Parameters: startRow, startColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 1, 4);

        // Set a formula in the merged cell to calculate a total.
        // Example: sum of values in column E from row 1 to row 10.
        sheet.Cells["A1"].Formula = "=SUM(E1:E10)";

        // Save the workbook (lifecycle: save)
        workbook.Save("MergedCellTotal.xlsx");
    }
}
