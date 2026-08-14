// Title: Freeze Top Row and First Two Columns in Excel with Aspose.Cells for .NET (C#)
// Description: Shows how to call worksheet.FreezePanes(1, 2, 1, 2) in Aspose.Cells for .NET to lock the first row and the first two columns, keeping headers visible while scrolling, and saves the file as FreezeTopRowFirstTwoColumns.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze top row Aspose.Cells | freeze first two columns Excel | lock header row Aspose.Cells | C# Excel freeze panes | Aspose.Cells worksheet freeze | freeze panes method parameters
// Common Searches: Aspose.Cells freeze top row C# example | how to freeze first two columns using Aspose.Cells | C# freeze panes Excel Aspose.Cells tutorial | freeze panes method Aspose.Cells documentation | keep header row visible Aspose.Cells workbook
// Developer Intent: Freeze the first row and the first two columns of a worksheet so header information remains visible during scrolling.
// Use Cases: Financial reports where the date row and account columns must stay in view. | Data‑entry templates that require the title row and identifier columns to be fixed. | Exported analytics dashboards where column headings and key IDs need constant visibility. | Large inventory sheets where product codes (first two columns) and headers are essential while navigating.
// AI Prompts: Generate C# code to freeze a custom range of rows and columns with Aspose.Cells. | Explain how the four parameters of FreezePanes map to split positions for variable worksheet sizes. | Provide an example that unfreezes panes, then refreezes a different set of rows and columns using Aspose.Cells. | Show how to apply FreezePanes to multiple worksheets in a single workbook programmatically.

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // Shows how to call worksheet.FreezePanes(1, 2, 1, 2) in Aspose.Cells for .NET to lock the first row and the first two columns, keeping headers visible while scrolling, and saves the file as FreezeTopRowFirstTwoColumns.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the top row (row index 0) and the first two columns (column indices 0 and 1)
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Row = 1 (second row) and Column = 2 (third column) define the split position.
            // freezedRows = 1 (freeze first row), freezedColumns = 2 (freeze first two columns)
            worksheet.FreezePanes(1, 2, 1, 2);

            // Save the workbook
            workbook.Save("FreezeTopRowFirstTwoColumns.xlsx");
        }
    }
}
