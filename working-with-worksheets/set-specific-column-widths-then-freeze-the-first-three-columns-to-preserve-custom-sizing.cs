// Title: C# – Set Column Widths and Freeze the First Three Columns with Aspose.Cells
// Description: Creates a new workbook, assigns custom character‑unit widths to columns A‑C, adds sample data, freezes those three columns using FreezePanes, and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# | set column width | freeze panes | first three columns | Excel column sizing | FreezePanes method | worksheet formatting | .NET Excel export
// Common Searches: Aspose.Cells set column width C# example | how to freeze first three columns with Aspose.Cells | C# FreezePanes after adjusting column widths | custom column widths Excel Aspose.Cells .NET | freeze panes and column sizing Aspose.Cells tutorial
// Developer Intent: Define explicit widths for columns A‑C and lock those columns in place so they remain visible during horizontal scrolling.
// Use Cases: Design a financial report where header columns need extra space and must stay visible while scrolling. | Build a data‑entry template with predefined column sizes and frozen panes to guide users. | Export large datasets to Excel with consistent column widths and frozen navigation columns for better readability.
// AI Prompts: Generate C# code using Aspose.Cells that sets column widths for A‑D and freezes the first two columns. | Show how to apply character‑unit column widths and simultaneously freeze rows and columns with Aspose.Cells. | Explain the mapping of FreezePanes parameters to row and column indices in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new workbook, assigns custom character‑unit widths to columns A‑C, adds sample data, freezes those three columns using FreezePanes, and saves the file as an XLSX document.
class SetColumnWidthsAndFreeze
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set specific column widths (character units)
        cells.SetColumnWidth(0, 20); // Column A
        cells.SetColumnWidth(1, 30); // Column B
        cells.SetColumnWidth(2, 15); // Column C

        // Sample data to illustrate column widths
        cells["A1"].PutValue("Column A data");
        cells["B1"].PutValue("Column B data with longer text");
        cells["C1"].PutValue("Col C");

        // Freeze the first three columns (A, B, C)
        // Freeze at column index 3 (D) with 0 frozen rows and 3 frozen columns
        sheet.FreezePanes(0, 3, 0, 3);

        // Save the workbook
        workbook.Save("ColumnWidthsAndFreeze.xlsx");
    }
}
