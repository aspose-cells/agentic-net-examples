// Title: How to set custom row heights and freeze the first two rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set row height for specific rows | How to freeze the first two rows in an Excel file using Aspose.Cells | Preserve custom row heights when freezing panes with Aspose.Cells C# | Aspose.Cells FreezePanes example for top rows | Set row heights and freeze panes in .xlsx using Aspose.Cells .NET
// Tags: row height customization Aspose.Cells | freeze top rows Aspose.Cells | FreezePanes usage Aspose.Cells C# | worksheet row formatting Aspose.Cells | preserve row height when freezing panes Aspose.Cells

using Aspose.Cells;
using System;

namespace AsposeCellsExample
{
    // The example creates a new workbook, sets custom heights for rows 1‑3, freezes the first two rows with the FreezePanes method, and saves the worksheet as CustomRowHeightsAndFreeze.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set custom row heights (height is in points)
                sheet.Cells.Rows[0].Height = 30; // Row 1 height
                sheet.Cells.Rows[1].Height = 40; // Row 2 height
                sheet.Cells.Rows[2].Height = 25; // Row 3 height

                // Freeze the first two rows (rows 0 and 1)
                // FreezePanes(row, column, totalRows, totalColumns)
                // row = 2 (first visible row after freeze), column = 0 (no column freeze)
                // totalRows = 2 (freeze two rows), totalColumns = 0 (no column freeze)
                sheet.FreezePanes(2, 0, 2, 0);

                // Save the workbook
                string outputFile = "CustomRowHeightsAndFreeze.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
