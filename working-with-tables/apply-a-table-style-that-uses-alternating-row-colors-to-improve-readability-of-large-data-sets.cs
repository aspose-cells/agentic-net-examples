// Title: C# – Apply Alternating Row Stripe Table Style with Aspose.Cells
// Description: Shows how to build a 100‑row × 5‑column worksheet, create a ListObject, set the built‑in TableStyleMedium2, turn on row banding, turn off column banding, and save the file as an XLSX workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# table style | row banding | Excel row stripes | ListObject styling | TableStyleMedium2 | disable column stripes | programmatic Excel formatting | large data set | alternating colors
// Common Searches: Aspose.Cells enable row stripes in a table | C# set built‑in table style with banded rows | how to hide column banding while keeping row banding in Excel via code | create ListObject with alternating row colors using Aspose.Cells | apply TableStyleMedium2 programmatically
// Developer Intent: The developer wants to programmatically add a table to a worksheet and use alternating row colors to make a large dataset easier to read.
// Use Cases: Produce a multi‑page report where the table rows are shaded alternately for quick visual scanning. | Export database query results to Excel with a predefined style that highlights rows but not columns. | Standardize the appearance of generated spreadsheets across an organization by applying a built‑in striped table style.
// AI Prompts: Generate C# code that toggles row and column stripe visibility for an Aspose.Cells ListObject. | Explain how to choose a different built‑in table style that includes row banding in Aspose.Cells for .NET. | Write a reusable method to apply a striped table style to any worksheet range using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableRowStripesDemo
{
    // Shows how to build a 100‑row × 5‑column worksheet, create a ListObject, set the built‑in TableStyleMedium2, turn on row banding, turn off column banding, and save the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large data set (e.g., 100 rows × 5 columns)
            // Add header row
            for (int col = 0; col < 5; col++)
            {
                cells[0, col].PutValue($"Header {col + 1}");
            }

            // Add data rows
            for (int row = 1; row <= 100; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row}C{col + 1}");
                }
            }

            // Create a ListObject (table) that covers the data range
            int firstRow = 0;
            int firstColumn = 0;
            int lastRow = 100;      // zero‑based index
            int lastColumn = 4;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style that supports row stripes
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Enable alternating row colors (row stripe formatting)
            table.ShowTableStyleRowStripes = true;

            // Optionally, disable column stripes if only row stripes are desired
            table.ShowTableStyleColumnStripes = false;

            // Save the workbook
            workbook.Save("TableWithRowStripes.xlsx", SaveFormat.Xlsx);
        }
    }
}
