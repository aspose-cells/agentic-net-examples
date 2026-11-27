// Title: Set custom page margins and freeze the first row in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to define custom left and right margins in points | C# freeze the top worksheet row with Aspose.Cells FreezePanes method | Set a print area that includes a frozen header row using Aspose.Cells | Configure page setup margins and freeze panes together in Aspose.Cells .NET
// Tags: custom margins page setup Aspose.Cells | freeze first row FreezePanes Aspose.Cells | worksheet print area definition Aspose.Cells | page layout configuration Aspose.Cells | Aspose.Cells C# margin and freeze example

using System;
using Aspose.Cells;

// The example creates a new workbook, applies 0.5‑inch left/right and 0.75‑inch top/bottom margins, freezes the first row with FreezePanes, sets the print area to include the frozen row, and saves the worksheet as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DemoSheet";

            // Fill some sample data so the effect of freezing can be observed
            for (int row = 0; row < 30; row++)
            {
                worksheet.Cells[row, 0].PutValue($"Item {row + 1}");
                worksheet.Cells[row, 1].PutValue((row + 1) * 10);
            }

            // -------------------------------------------------
            // 1. Define custom page margins (values are in points; 1 inch = 72 points)
            // -------------------------------------------------
            // Left and right margins: 0.5 inch
            // Top and bottom margins: 0.75 inch
            worksheet.PageSetup.LeftMargin   = 0.5 * 72; // 36 points
            worksheet.PageSetup.RightMargin  = 0.5 * 72; // 36 points
            worksheet.PageSetup.TopMargin    = 0.75 * 72; // 54 points
            worksheet.PageSetup.BottomMargin = 0.75 * 72; // 54 points

            // -------------------------------------------------
            // 2. Freeze the top row so it stays visible (and within the printable area)
            // -------------------------------------------------
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Freeze the first row (index 0) – we pass row=1 because the method expects the
            // cell just below the frozen pane. No columns are frozen.
            worksheet.FreezePanes(1, 0, 1, 0);

            // Optional: ensure the frozen row is part of the print area
            worksheet.PageSetup.PrintArea = "A1:B30";

            // Save the workbook
            string outputPath = "CustomMarginAndFreeze.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
