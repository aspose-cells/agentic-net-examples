// Title: Auto‑fit all rows then freeze the top row in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells auto fit rows before applying FreezePanes | how to adjust row heights and freeze first row with Aspose.Cells .NET | Aspose.Cells example to auto‑fit rows and freeze pane at A2 | auto fit rows then freeze top row in Excel using Aspose.Cells C# code | preserve row height consistency when freezing panes with Aspose.Cells
// Tags: AutoFitRows method Aspose.Cells C# | FreezePanes API after row auto‑fit | maintain row height consistency Aspose.Cells | worksheet row height adjustment Aspose.Cells | freeze first row Excel Aspose.Cells .NET

using Aspose.Cells;

// The sample loads an Excel workbook, uses Aspose.Cells to auto‑fit the heights of all rows on the first worksheet, freezes the top row with FreezePanes, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Auto‑fit all rows to adjust their heights based on content
        sheet.AutoFitRows();

        // Freeze the first row (you can adjust parameters to freeze more rows/columns)
        // FreezePanes(totalRows, totalColumns, rows, columns)
        // Here we freeze 1 row, 0 columns; the scrollable area starts at cell A2 (row index 1)
        sheet.FreezePanes(1, 0, 1, 0);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
