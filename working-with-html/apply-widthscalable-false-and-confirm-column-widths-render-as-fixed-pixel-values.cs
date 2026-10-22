// Title: How to set a fixed pixel width for a column in Aspose.Cells for .NET and confirm the value
// AI Prompts: Write C# code using Aspose.Cells to set column A to 120 pixels with SetColumnWidthPixel, then read the width using GetColumnWidthPixel and print it to the console. | Demonstrate creating a new workbook, applying a non‑scaled column width, and saving the file as WidthScalableFalse.xlsx. | Show how to verify that a column width is stored as a fixed pixel value without relying on the deprecated IsColumnWidthScaled property.
// Common Searches: Aspose.Cells C# set column width in pixels and read back value | How to disable column width scaling in Aspose.Cells .NET | Get fixed pixel column width from worksheet using Aspose.Cells | Save Excel file with exact column pixel width using Aspose.Cells | Verify column width is not auto‑scaled in Aspose.Cells workbook
// Tags: pixel column width assignment Aspose.Cells C# | retrieve fixed column width pixel Aspose.Cells | column width scaling off Aspose.Cells | save workbook with exact column pixel width Aspose.Cells | confirm column width value Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, sets column A's width to 120 pixels via SetColumnWidthPixel, reads the width back with GetColumnWidthPixel to confirm it is stored as a fixed pixel value, prints the result, ensures the output directory exists, saves the workbook as WidthScalableFalse.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: In recent Aspose.Cells versions the IsColumnWidthScaled property
            // is not available. Fixed pixel widths can be set directly via
            // SetColumnWidthPixel, so we omit the scaling setting.

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the width of column A (index 0) to a fixed pixel value, e.g., 120 pixels
            sheet.Cells.SetColumnWidthPixel(0, 120);

            // Retrieve the pixel width to confirm it is stored as a fixed value
            double pixelWidth = sheet.Cells.GetColumnWidthPixel(0);
            Console.WriteLine($"Column A width (fixed pixels): {pixelWidth}");

            // Prepare output path and ensure the directory exists
            string outputPath = "WidthScalableFalse.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
