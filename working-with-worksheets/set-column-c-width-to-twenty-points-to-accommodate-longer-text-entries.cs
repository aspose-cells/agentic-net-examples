// Title: Set column C width to 20 points (pixel conversion) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column width in points | convert points to pixels for Excel column width with Aspose.Cells | set column C width to 20 points using SetColumnWidthPixel | adjust Excel column width by points in a .NET application | calculate pixel value from points for column sizing in Aspose.Cells
// Tags: set column width pixel Aspose.Cells | column width points to pixels .NET | Aspose.Cells column C width adjustment | Excel column sizing SetColumnWidthPixel | C# Aspose.Cells workbook column width

using System;
using Aspose.Cells;

namespace SetColumnWidthExample
{
    // The example creates (or loads) a workbook, accesses the first worksheet, converts 20 points to pixels based on a 96 dpi screen, and applies the resulting pixel width to column C (index 2) using SetColumnWidthPixel. Sample long text is written to cells C1 and C2, and the workbook is saved as ColumnC_Width20Points.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook (or load an existing one)
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook(); // creates a blank workbook
                // If you have an existing file, uncomment the line below and provide the path:
                // Workbook workbook = new Workbook("input.xlsx");

                // -----------------------------------------------------------------
                // 2. Access the first worksheet
                // -----------------------------------------------------------------
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 3. Convert 20 points to pixels.
                //    1 point = 1/72 inch, typical screen DPI = 96 pixels per inch.
                // -----------------------------------------------------------------
                const double points = 20.0;
                double inches = points / 72.0;               // inches
                int pixels = (int)Math.Round(inches * 96);   // pixels (rounded)

                // -----------------------------------------------------------------
                // 4. Set column C (zero‑based index 2) width to the calculated pixels
                // -----------------------------------------------------------------
                // The Cells collection provides SetColumnWidthPixel for direct pixel sizing.
                worksheet.Cells.SetColumnWidthPixel(2, pixels);

                // -----------------------------------------------------------------
                // 5. (Optional) Write some sample data to demonstrate the width
                // -----------------------------------------------------------------
                worksheet.Cells["C1"].PutValue("This is a longer text that should fit.");
                worksheet.Cells["C2"].PutValue("Another long entry.");

                // -----------------------------------------------------------------
                // 6. Save the workbook to disk
                // -----------------------------------------------------------------
                string outputPath = "ColumnC_Width20Points.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Column C width set to {points} points ({pixels} pixels).");
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Any exception will be printed to the console – no hidden failures.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
