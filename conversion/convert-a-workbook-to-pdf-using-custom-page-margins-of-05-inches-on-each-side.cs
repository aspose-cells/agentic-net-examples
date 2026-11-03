// Title: Convert Excel workbook to PDF with 0.5‑inch page margins in C# using Aspose.Cells
// Common Searches: c# aspnet set page margins to 0.5 inch before exporting Excel to PDF with Aspose.Cells | aspocells convert xlsx to pdf with custom 0.5 inch margins | how to apply the same margins to all worksheets when saving as PDF in Aspose.Cells | set left right top bottom margins in points for PDF output using Aspose.Cells C# | handle missing file error while converting Excel to PDF with Aspose.Cells
// Tags: Aspose.Cells worksheet page setup margins | custom margins Excel to PDF conversion C# | uniform page margins Aspose.Cells PDF export | 0.5 inch margin points configuration Aspose.Cells | SaveFormat.Pdf margin settings Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an .xlsx file, sets each worksheet's left, right, top, and bottom margins to 0.5 inches (36 points) via PageSetup, and saves the workbook as a PDF using Aspose.Cells, including checks for missing files and CellsException handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Define margin size: 0.5 inches = 36 points (1 inch = 72 points)
            const double marginPoints = 36.0;

            // Apply the custom margins to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.LeftMargin = marginPoints;
                pageSetup.RightMargin = marginPoints;
                pageSetup.TopMargin = marginPoints;
                pageSetup.BottomMargin = marginPoints;
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF file saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handle password-protected files or other Cells-specific errors
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
