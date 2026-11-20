// Title: Adjust Excel column widths by pixel using Aspose.Cells AutoFit and SetColumnWidthPixel in C#
// Common Searches: C# Aspose.Cells set column width in pixels after autofit | increase Excel column width by a few pixels using GetColumnWidthPixel | programmatically add padding to column widths in Aspose.Cells .NET | measure column width in pixels with Aspose.Cells and apply custom width | auto‑fit columns then adjust pixel width in Aspose.Cells workbook
// Tags: auto-fit column Aspose.Cells | SetColumnWidthPixel method | GetColumnWidthPixel method | pixel‑based column sizing | column width padding technique

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, iterates through each worksheet and its columns, auto‑fits each column, reads the resulting pixel width with GetColumnWidthPixel, adds a small padding (e.g., 5 px), and then sets the new width using SetColumnWidthPixel before saving the file.
class AdjustColumnWidthByPixel
{
    static void Main()
    {
        const string inputPath = @"Input.xlsx";
        const string outputPath = @"Output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the maximum column index that contains data
                int maxColumn = cells.MaxColumn;

                // Iterate over each column that may contain data
                for (int col = 0; col <= maxColumn; col++)
                {
                    // Auto‑fit the column based on its content
                    sheet.AutoFitColumn(col);

                    // Retrieve the auto‑fitted width in pixels
                    int currentPixelWidth = cells.GetColumnWidthPixel(col);

                    // Add a small padding (e.g., 5 pixels) to avoid clipping
                    int finalPixelWidth = currentPixelWidth + 5;

                    // Set the column width in pixels
                    cells.SetColumnWidthPixel(col, finalPixelWidth);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
