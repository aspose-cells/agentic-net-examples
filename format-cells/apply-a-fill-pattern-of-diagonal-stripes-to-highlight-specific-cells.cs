using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class DiagonalStripeHighlight
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range to be highlighted (e.g., B2:D4)
            int startRow = 1;      // B2 (zero‑based)
            int startColumn = 1;   // Column B
            int totalRows = 3;     // Rows 2‑4 inclusive
            int totalColumns = 3;  // Columns B‑D inclusive

            // Create a style with a diagonal stripe pattern
            Style stripeStyle = workbook.CreateStyle();
            stripeStyle.Pattern = BackgroundType.DiagonalStripe;
            stripeStyle.ForegroundColor = Color.Yellow;
            stripeStyle.BackgroundColor = Color.Blue;

            // Enable cell shading so the pattern is applied
            StyleFlag flag = new StyleFlag { CellShading = true };

            // Apply the style to the defined range
            AsposeRange range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
            range.ApplyStyle(stripeStyle, flag);

            // Save the workbook
            workbook.Save("DiagonalStripeHighlight.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}