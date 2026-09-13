// Title: How to merge cells A1:D1 and set a SUM formula in Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, merge the range A1:D1 on the first worksheet, assign the formula =SUM(B2:D2) to the merged cell, and save the file as MergedCellFormula.xlsx using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically merge cells across columns A to D in row 1 and set a SUM formula that references B2 through D2. | Write C# code that opens a workbook, merges cells A1 through D1, applies a SUM formula to the merged cell, and persists the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add SUM formula after merging cells | programmatically set formula for merged cells in Aspose.Cells | example of merging cells and calculating total with Aspose.Cells .NET | C# Aspose.Cells merge cells and assign formula to merged range | calculate sum of B2:D2 in merged cell using Aspose.Cells
// Tags: Aspose.Cells merge first row cells C# | assign aggregate formula to merged cell Aspose.Cells | create workbook programmatically Aspose.Cells .NET | apply cell formula across range Aspose.Cells | save workbook as XLSX Aspose.Cells

using Aspose.Cells;

// The sample creates a new workbook, merges cells A1:D1 on the first worksheet, sets a SUM formula that totals B2:D2 in the merged cell, and saves the workbook as MergedCellFormula.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided creation rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells from column A to D in the first row (A1:D1)
        // Parameters: startRow, startColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 1, 4);

        // Set a formula in the merged cell.
        // Example: calculate the sum of values in B2:D2
        sheet.Cells["A1"].Formula = "=SUM(B2:D2)";

        // Save the workbook (using the provided save rule)
        workbook.Save("MergedCellFormula.xlsx");
    }
}
