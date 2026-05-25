using System;
using System.IO;
using Aspose.Cells;

namespace DynamicColumnWidthDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data with varying lengths
                cells["A1"].PutValue("Short");
                cells["A2"].PutValue("A much longer piece of text that should expand the column");
                cells["B1"].PutValue(12345);
                cells["B2"].PutValue("Medium length");
                cells["C1"].PutValue("Tiny");
                cells["C2"].PutValue("Another very long text string that will require a wider column");

                // Determine the used range to know how many columns to inspect
                int maxColumn = cells.MaxDataColumn;
                int maxRow = cells.MaxDataRow; // needed for AutoFitColumn overload

                // Padding to add a little extra space (in pixels)
                const int paddingPixels = 5;

                // Iterate through each column, auto‑fit, then add padding
                for (int col = 0; col <= maxColumn; col++)
                {
                    // Auto‑fit the column based on its content (firstRow = 0, lastRow = maxRow)
                    sheet.AutoFitColumn(col, 0, maxRow);

                    // Get the current width in pixels, add padding, and set it back
                    int currentWidth = cells.GetColumnWidthPixel(col);
                    int finalWidth = currentWidth + paddingPixels;
                    cells.SetColumnWidthPixel(col, finalWidth);
                }

                // Prepare output path
                string outputPath = "DynamicColumnWidthDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}