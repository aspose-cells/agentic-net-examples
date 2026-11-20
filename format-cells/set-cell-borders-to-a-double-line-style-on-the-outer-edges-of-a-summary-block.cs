// Title: How to apply double‑line outer borders to a summary block range in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set double line border on a range of cells | Apply only outer borders to a block of cells in Excel with Aspose.Cells .NET | Create black double‑border style for a summary section using Aspose.Cells API | Use StyleFlag to format borders without affecting cell values in Aspose.Cells C# | Set double line top, bottom, left, and right borders for a specific range in a .NET Excel library
// Tags: outer cell border styling Aspose.Cells | range border style Aspose.Cells .NET | StyleFlag borders only | black cell border formatting | summary block formatting Excel

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, defines a range covering rows 5‑10 and columns B‑E, builds a style with black double‑line borders on the top, bottom, left, and right edges, applies the style to the range using a StyleFlag that targets borders only, and saves the result as output.xlsx.
class SetSummaryBlockBorders
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the summary block range (example: rows 5-10, columns B-E)
            int startRow = 4;      // zero‑based index (row 5)
            int startColumn = 1;   // zero‑based index (column B)
            int endRow = 9;        // row 10
            int endColumn = 4;     // column E

            // Create the range representing the summary block
            int totalRows = endRow - startRow + 1;
            int totalColumns = endColumn - startColumn + 1;
            Aspose.Cells.Range summaryRange = worksheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Create a new style for the range
            Style rangeStyle = workbook.CreateStyle();

            // Set double line borders on the outer edges
            rangeStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
            rangeStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
            rangeStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
            rangeStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;

            // Optionally set border color (e.g., black)
            rangeStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            rangeStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
            rangeStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            rangeStyle.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply the style to the range (only borders)
            summaryRange.ApplyStyle(rangeStyle, new StyleFlag { Borders = true });

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
