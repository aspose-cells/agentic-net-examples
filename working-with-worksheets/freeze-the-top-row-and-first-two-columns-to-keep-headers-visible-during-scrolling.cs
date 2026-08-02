// Title: Freeze Top Row and First Two Columns in Excel with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills it with sample data, and calls Worksheet.FreezePanes(1, 2, 1, 2) to lock the first row and the first two columns, then saves the file as FreezeTopRowAndFirstTwoColumns.xlsx.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze top row | freeze first two columns | Excel pane freezing | worksheet freeze panes .NET
// Common Searches: Aspose.Cells freeze top row C# | How to freeze first two columns using Aspose.Cells | Worksheet.FreezePanes example .NET | Freeze panes programmatically in Excel with Aspose | C# code to lock header row and columns in generated spreadsheet
// Developer Intent: The developer wants header rows and leading identifier columns to remain visible while scrolling through large worksheets generated with Aspose.Cells.
// Use Cases: Generating financial reports where column headers and row identifiers must stay fixed during navigation. | Exporting data grids from web applications with a summary row at the top and key columns on the left for better readability. | Creating large inventory sheets where frozen panes improve user experience when scrolling through thousands of rows and columns.
// AI Prompts: Write C# code using Aspose.Cells to freeze the first row and the first two columns of a worksheet. | Explain the parameters of Worksheet.FreezePanes and demonstrate how to adjust them for different pane configurations. | Create a reusable function that accepts row and column counts and applies FreezePanes accordingly in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // Creates a workbook, fills it with sample data, and calls Worksheet.FreezePanes(1, 2, 1, 2) to lock the first row and the first two columns, then saves the file as FreezeTopRowAndFirstTwoColumns.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to demonstrate scrolling
            for (int i = 0; i < 20; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    worksheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Freeze the top row and the first two columns
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            // Row index = 1 (freeze rows above it), Column index = 2 (freeze columns left of it)
            worksheet.FreezePanes(1, 2, 1, 2);

            // Save the workbook
            workbook.Save("FreezeTopRowAndFirstTwoColumns.xlsx");
        }
    }
}
