// Title: C# Aspose.Cells – Conditionally Freeze Panes on Sheets with Over 10 Columns
// Description: The sample builds a workbook, adds two worksheets with different numbers of populated columns, reads the column count via MaxDataColumn, and invokes FreezePanes(1,1,1,1) only for sheets that exceed ten columns, then saves the workbook.
// Keywords: aspose.cells | c# | .net | freeze panes | conditional freeze | maxdatacolumn | column count detection | worksheet automation | excel export | freeze first row and column
// Common Searches: aspose.cells freeze panes based on column count | c# check number of columns before freezing panes | apply FreezePanes only when sheet has more than 10 columns | use MaxDataColumn to conditionally freeze rows in Aspose.Cells | conditional freeze panes example in .NET
// Developer Intent: Apply FreezePanes(1,1,1,1) to any worksheet that contains more than ten data columns.
// Use Cases: Generating multi‑sheet reports where wide tables automatically lock the header row and left column for easier navigation. | Building a template that freezes panes during data import only on sheets exceeding ten columns, ensuring consistent view settings. | Processing a batch of worksheets, detecting those with extensive column data, and selectively applying FreezePanes before distribution.
// AI Prompts: Create C# code using Aspose.Cells that iterates through all worksheets, counts populated columns with MaxDataColumn, and calls FreezePanes(1,1,1,1) only when the count is greater than 10. | Show an example that adds sample data to two sheets, determines column count, conditionally freezes the top row and first column, and saves the file as an .xlsx. | Explain how MaxDataColumn can be used to decide whether to apply FreezePanes in Aspose.Cells, and provide the corresponding C# implementation.

using System;
using Aspose.Cells;

namespace FreezeIfMoreThanTenColumnsDemo
{
    // The sample builds a workbook, adds two worksheets with different numbers of populated columns, reads the column count via MaxDataColumn, and invokes FreezePanes(1,1,1,1) only for sheets that exceed ten columns, then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Example: add some data to demonstrate column count detection
            Worksheet sheet1 = workbook.Worksheets[0];
            for (int col = 0; col < 12; col++) // 12 columns > 10
            {
                sheet1.Cells[0, col].PutValue($"Header{col + 1}");
            }

            Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            for (int col = 0; col < 8; col++) // 8 columns <= 10
            {
                sheet2.Cells[0, col].PutValue($"Header{col + 1}");
            }

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Determine the last column index that contains data.
                // MaxDataColumn returns zero‑based index; add 1 to get count.
                int columnCount = ws.Cells.MaxDataColumn + 1;

                // Apply freeze panes only if the worksheet has more than ten columns.
                if (columnCount > 10)
                {
                    // Freeze the first row and first column (cell B2) – 1 row and 1 column frozen.
                    ws.FreezePanes(1, 1, 1, 1);
                }
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FreezeIfMoreThan10Columns.xlsx");
        }
    }
}
