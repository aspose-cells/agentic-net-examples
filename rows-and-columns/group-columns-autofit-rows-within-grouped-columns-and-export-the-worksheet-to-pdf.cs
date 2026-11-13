// Title: Group columns B‑D, auto‑fit rows, and save the worksheet as a PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group specific columns and auto‑fit rows before PDF export | C# code to collapse columns B to D and adjust row heights with Aspose.Cells | How to generate a PDF from a worksheet after grouping columns in Aspose.Cells .NET | AutoFitRows after GroupColumns example using Aspose.Cells for .NET
// Tags: group columns Aspose.Cells C# | auto‑fit rows Aspose.Cells | export worksheet to PDF Aspose.Cells | column grouping PDF conversion .NET | adjust row height after column grouping Aspose.Cells

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, filling sample data, grouping columns B‑D, auto‑fitting all rows to show grouped content, and saving the worksheet as a PDF file with Aspose.Cells for .NET.
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

        // Group columns B to D (column indexes 1, 2, 3)
        // Parameters: start column index, total number of columns to group
        sheet.Cells.GroupColumns(1, 3);

        // Auto‑fit all rows so that content within the grouped columns is fully visible
        sheet.AutoFitRows();

        // Export the worksheet to a PDF file
        workbook.Save("GroupedColumns.pdf", SaveFormat.Pdf);
    }
}
