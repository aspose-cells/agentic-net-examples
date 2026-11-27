// Title: How to apply the Light1 theme background to alternate rows in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells apply Light1 theme color to every other row in a worksheet | Create zebra striping with Light1 background in Excel using Aspose.Cells .NET | Apply alternating row style with StyleFlag in Aspose.Cells C# example | Set solid light gray fill for even rows in an Excel file with Aspose.Cells | How to use CreateRange and ApplyStyle for alternate rows in Aspose.Cells
// Tags: Light1 theme background formatting Aspose.Cells | alternating row shading Excel .NET | StyleFlag usage for range styling Aspose.Cells | CreateRange method for row selection Aspose.Cells | zebra striping implementation Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, fills a 10‑row by 4‑column table, defines a light gray style as a fallback for the Light1 theme, and uses CreateRange with a StyleFlag to apply the style to every even (1‑based) row, producing an alternating‑row background effect before saving the file as AlternateRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data (optional)
            for (int r = 0; r < 10; r++)
            {
                for (int c = 0; c < 4; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }

            // Define a style that uses a light background color (approximation of theme's Light1)
            var style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGray;   // fallback to a light gray color
            style.Pattern = BackgroundType.Solid;      // Solid fill

            // Apply the style to alternate rows (e.g., rows 2,4,6,...)
            int startRow = 0;      // zero‑based index of the first row of the table
            int totalRows = 10;    // total number of rows in the table
            int totalCols = 4;     // total number of columns in the table

            for (int row = startRow; row < startRow + totalRows; row++)
            {
                // Apply to even rows (1‑based even => zero‑based odd)
                if ((row - startRow) % 2 == 1)
                {
                    var range = sheet.Cells.CreateRange(row, 0, 1, totalCols);
                    var flag = new StyleFlag { All = true };
                    range.ApplyStyle(style, flag);
                }
            }

            // Determine output file path and ensure the directory exists
            string outputPath = "AlternateRows.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
