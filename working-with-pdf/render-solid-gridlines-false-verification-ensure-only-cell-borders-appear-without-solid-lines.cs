// Title: How to hide Excel gridlines and show only cell borders with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide worksheet gridlines but keep custom borders | Check IsGridlinesVisible property after disabling gridlines in Aspose.Cells | Apply thin borders to a range while removing default gridlines using Aspose.Cells .NET | Save Excel workbook without solid gridlines using Aspose.Cells C# | Verify that only cell borders are visible in generated Excel file with Aspose.Cells
// Tags: disable worksheet gridlines Aspose.Cells | apply thin borders to range Aspose.Cells | gridlines visibility verification Aspose.Cells | export Excel without solid gridlines .NET | styleflag borders only Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsGridlinesDemo
{
    // The example creates a new workbook, turns off the default worksheet gridlines, fills a 3x3 range with data, applies thin borders to the range using a Style and StyleFlag, confirms that IsGridlinesVisible is false, and saves the file as GridlinesVerification.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Disable solid gridlines so that only cell borders are visible
                // ------------------------------------------------------------
                sheet.IsGridlinesVisible = false; // Hide default gridlines

                // ------------------------------------------------------------
                // 2. Add sample data and apply borders to demonstrate the effect
                // ------------------------------------------------------------
                // Populate a 3x3 range with sample values
                for (int row = 0; row < 3; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Apply a thin border to the entire range A1:C3
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:C3");
                Style style = workbook.CreateStyle();
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                // Apply the style to the range (only borders are affected)
                range.ApplyStyle(style, new StyleFlag { Borders = true });

                // ------------------------------------------------------------
                // 3. Verification: ensure that solid gridlines are disabled
                // ------------------------------------------------------------
                bool gridlinesVisible = sheet.IsGridlinesVisible;
                Console.WriteLine($"Solid gridlines visible: {gridlinesVisible} (expected: False)");

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                string outputPath = Path.Combine(Environment.CurrentDirectory, "GridlinesVerification.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
