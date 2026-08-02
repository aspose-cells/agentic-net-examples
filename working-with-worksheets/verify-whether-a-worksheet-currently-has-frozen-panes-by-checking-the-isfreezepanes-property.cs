// Title: Detect Frozen Panes in a Worksheet with Aspose.Cells C# (IsFreezePanes/GetFreezedPanes)
// Description: This C# example creates a workbook, applies FreezePanes at cell C5, then uses GetFreezedPanes (or the IsFreezePanes property) to verify whether the worksheet contains frozen panes and to retrieve the freeze position, frozen rows, and frozen columns before saving the file.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | GetFreezedPanes | IsFreezePanes | detect frozen panes | worksheet freeze status | retrieve frozen rows columns
// Common Searches: Aspose.Cells check if worksheet has frozen panes C# | GetFreezedPanes example Aspose.Cells | IsFreezePanes property usage | how to read freeze pane coordinates Aspose.Cells | detect frozen rows and columns in .NET workbook
// Developer Intent: Find out whether a worksheet currently has frozen panes and obtain the exact freeze coordinates.
// Use Cases: Validate that a generated report keeps header rows frozen before distribution. | Conditionally apply formatting only when no panes are frozen to prevent layout issues. | Log freeze pane details for audit trails when exporting workbooks to clients.
// AI Prompts: Generate C# code using Aspose.Cells to determine if a worksheet has frozen panes and print the freeze row and column. | Show how to call GetFreezedPanes after FreezePanes to retrieve frozen rows and columns. | Explain when to prefer IsFreezePanes versus GetFreezedPanes in Aspose.Cells.

using Aspose.Cells;
using System;

// This C# example creates a workbook, applies FreezePanes at cell C5, then uses GetFreezedPanes (or the IsFreezePanes property) to verify whether the worksheet contains frozen panes and to retrieve the freeze position, frozen rows, and frozen columns before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze panes at cell C5 (row index 4, column index 2) with 2 frozen rows and 1 frozen column
        worksheet.FreezePanes(4, 2, 2, 1);

        // Variables to receive freeze pane details
        int row, column, frozenRows, frozenColumns;

        // GetFreezedPanes returns true if the worksheet has frozen panes
        bool hasFreeze = worksheet.GetFreezedPanes(out row, out column, out frozenRows, out frozenColumns);

        // Display the result
        Console.WriteLine("Worksheet has frozen panes: " + hasFreeze);
        if (hasFreeze)
        {
            Console.WriteLine($"Freeze position - Row: {row}, Column: {column}");
            Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
        }

        // Save the workbook (optional)
        workbook.Save("FrozenPanesCheck.xlsx");
    }
}
