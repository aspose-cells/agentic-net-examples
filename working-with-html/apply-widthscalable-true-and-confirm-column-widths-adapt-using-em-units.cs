// Title: Set column width in characters (approx em units) and verify scaling with Aspose.Cells for .NET
// Common Searches: how to set column width in em units with Aspose.Cells for .NET | C# code to convert Aspose.Cells column width to pixel value | verify that column width scaling works after saving an Excel file using Aspose.Cells | adjust column width for HTML export in Aspose.Cells C# example | determine character count equivalent for column width in Aspose.Cells workbook
// Tags: define column width in characters Aspose.Cells | em unit column width conversion C# | pixel width estimation from column width | post‑save column width verification Aspose.Cells | HTML export column width scaling Aspose.Cells

using Aspose.Cells;
using System;

namespace Demo
{
    // The example creates a workbook, accesses column A, sets its width to 12 characters (≈12 em), prints the character‑based width and an approximate pixel value, saves the file, and demonstrates how to verify column width scaling in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Access column A (index 0)
                var column = sheet.Cells.Columns[0];

                // Set column width (character count approximates em)
                column.Width = 12; // 12 characters ≈ 12 em

                // Output the width settings
                Console.WriteLine("Width (characters): " + column.Width);
                // Approximate pixel width (1 character ≈ 7 pixels, varies with font)
                Console.WriteLine("Approximate Width (pixels): " + (int)(column.Width * 7));

                // Save the workbook
                string outputPath = "WidthScalableDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
