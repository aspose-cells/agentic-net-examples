// Title: Create a double‑line outer border around a summary block range in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set double line border on the outer edges of a cell range | How to apply thick double borders around a summary block in Excel with Aspose.Cells .NET | Aspose.Cells example for adding top, bottom, left and right double borders to a range | Apply double border style to a specific range (B2:D5) using Aspose.Cells in C# | Set outer double borders for a table area in an Excel workbook with Aspose.Cells API
// Tags: double line border styling Aspose.Cells | outer edge border configuration C# | range border application Excel .NET | styleflag usage for cell borders Aspose | create range and assign border style Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, defines a summary block range (B2:D5), and applies a double‑line border to the top, bottom, left, and right edges by configuring separate Style and StyleFlag objects for each side, then saves the file as SummaryBlock.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the summary block range (example: B2:D5)
                int startRow = 1;      // B2 -> row index 1 (zero‑based)
                int startColumn = 1;   // B2 -> column index 1
                int endRow = 4;        // D5 -> row index 4
                int endColumn = 3;     // D5 -> column index 3

                int totalRows = endRow - startRow + 1;
                int totalColumns = endColumn - startColumn + 1;

                // ---------- Top border ----------
                Style topStyle = workbook.CreateStyle();
                topStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
                StyleFlag topFlag = new StyleFlag { Borders = true, TopBorder = true };
                AsposeRange topRange = sheet.Cells.CreateRange(startRow, startColumn, 1, totalColumns);
                topRange.ApplyStyle(topStyle, topFlag);

                // ---------- Bottom border ----------
                Style bottomStyle = workbook.CreateStyle();
                bottomStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
                StyleFlag bottomFlag = new StyleFlag { Borders = true, BottomBorder = true };
                AsposeRange bottomRange = sheet.Cells.CreateRange(endRow, startColumn, 1, totalColumns);
                bottomRange.ApplyStyle(bottomStyle, bottomFlag);

                // ---------- Left border ----------
                Style leftStyle = workbook.CreateStyle();
                leftStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
                StyleFlag leftFlag = new StyleFlag { Borders = true, LeftBorder = true };
                AsposeRange leftRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, 1);
                leftRange.ApplyStyle(leftStyle, leftFlag);

                // ---------- Right border ----------
                Style rightStyle = workbook.CreateStyle();
                rightStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                StyleFlag rightFlag = new StyleFlag { Borders = true, RightBorder = true };
                AsposeRange rightRange = sheet.Cells.CreateRange(startRow, endColumn, totalRows, 1);
                rightRange.ApplyStyle(rightStyle, rightFlag);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "SummaryBlock.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
