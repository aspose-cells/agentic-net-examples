using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertLogoAndFreeze
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the logo image file
            string logoPath = "logo.png";

            // Verify that the logo file exists before attempting to add it
            if (File.Exists(logoPath))
            {
                // Insert the logo picture anchored to cell A1 (row 0, column 0)
                int picIndex = sheet.Pictures.Add(0, 0, logoPath);
                Picture logo = sheet.Pictures[picIndex];

                // Place the picture inside the cell so it moves with the row/column
                logo.IsPlacedInCell = true;
            }
            else
            {
                Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping logo insertion.");
            }

            // Freeze the top rows that contain the logo (e.g., first 5 rows)
            int frozenRows = 5; // number of rows to keep visible
            sheet.FreezePanes(frozenRows, 0, frozenRows, 0);

            // Save the workbook
            string outputPath = "LogoWithFrozenRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}