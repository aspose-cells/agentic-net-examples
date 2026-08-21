// Title: Freeze the first two columns in Excel using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate a 20‑row by 10‑column range, and call Worksheet.FreezePanes(0, 2, 0, 2) to lock columns A and B while allowing horizontal scrolling. The workbook is saved as FreezeFirstTwoColumns.xlsx.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze columns Excel | static identifier columns | horizontal scrolling | worksheet.FreezePanes example | Excel .NET API | freeze first two columns | Aspose.Cells tutorial
// Common Searches: Aspose.Cells freeze first two columns C# | How to lock columns A and B with FreezePanes in .NET | Freeze panes Excel using Aspose.Cells API | C# code to freeze columns in an Excel workbook | Aspose.Cells example for freezing columns
// Developer Intent: Keep the first two columns visible while scrolling horizontally in an Excel worksheet generated with Aspose.Cells.
// Use Cases: Financial or inventory reports where ID and code columns must stay in view. | Data‑entry templates that require reference columns to remain static. | Exporting large tables to Excel with frozen identifier columns for easier navigation.
// AI Prompts: Generate C# code with Aspose.Cells that freezes the first three columns of a worksheet and saves the file. | Show how to freeze both a header row and the first two columns simultaneously using FreezePanes in Aspose.Cells for .NET. | Provide an Aspose.Cells example that freezes columns A‑B, adds a frozen header row, and formats the header.

using System;
using Aspose.Cells;

namespace FreezeFirstTwoColumnsDemo
{
    // Demonstrates how to create a workbook, populate a 20‑row by 10‑column range, and call Worksheet.FreezePanes(0, 2, 0, 2) to lock columns A and B while allowing horizontal scrolling. The workbook is saved as FreezeFirstTwoColumns.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to visualize the freeze effect
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first two columns (A and B)
            // Freeze at column index 2 (C) with 0 frozen rows and 2 frozen columns
            worksheet.FreezePanes(0, 2, 0, 2);

            // Save the workbook
            workbook.Save("FreezeFirstTwoColumns.xlsx");
        }
    }
}
