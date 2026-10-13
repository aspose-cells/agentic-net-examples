// Title: Group columns B‑D, auto‑fit rows, and save worksheet as PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to collapse columns B through D, auto‑fit every row, and generate a PDF file of the worksheet. | Create a .NET workbook, apply column grouping with collapse, adjust row heights automatically, then export the sheet to a PDF using Aspose.Cells.
// Common Searches: aspnet how to collapse a range of columns and export to PDF with Aspose.Cells | c# auto fit rows after grouping columns in Aspose.Cells | save worksheet as PDF after grouping columns B to D using Aspose.Cells | Aspose.Cells .NET group columns and auto fit rows before PDF conversion
// Tags: group columns collapse Aspose.Cells .NET | auto-fit rows Aspose.Cells | export worksheet to PDF Aspose.Cells | column grouping with collapse C# | pdf conversion after column grouping Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills it with data, collapses columns B‑D, auto‑fits all rows, and saves the worksheet as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Group columns B to D (column indexes 1 to 3)
        // The third parameter 'true' indicates that the grouped columns will be collapsed
        sheet.Cells.GroupColumns(1, 3, true);

        // Auto‑fit all rows so that row heights adjust to the content,
        // including rows that intersect the grouped columns
        sheet.AutoFitRows(0, sheet.Cells.MaxDataRow);

        // Export the worksheet to a PDF file
        workbook.Save("GroupedColumns.pdf", SaveFormat.Pdf);
    }
}
