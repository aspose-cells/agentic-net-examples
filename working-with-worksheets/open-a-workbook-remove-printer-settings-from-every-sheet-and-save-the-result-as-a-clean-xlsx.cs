// Title: How to remove all printer and page‑setup settings from each worksheet in an XLSX file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear printer settings for all worksheets in an existing workbook | remove print area and margins from every sheet with Aspose.Cells .NET | reset page setup options for all worksheets and save as new XLSX using Aspose.Cells | C# code to delete print titles and fit‑to‑page settings in an Excel file with Aspose.Cells | Aspose.Cells remove page‑setup flags before exporting workbook
// Tags: clear worksheet printer settings Aspose.Cells | reset page setup margins C# | remove print area Aspose.Cells workbook | save cleaned Excel file Aspose.Cells | iterate worksheets page setup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX workbook, iterates through every worksheet to clear print area, title rows/columns, margins, fit‑to‑page and other page‑setup options, then saves the modified workbook as a new clean XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "clean.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear printer‑related settings
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear print area and title rows/columns
                sheet.PageSetup.PrintArea = "";
                sheet.PageSetup.PrintTitleRows = "";
                sheet.PageSetup.PrintTitleColumns = "";

                // Reset margins to default (0) – adjust as needed
                sheet.PageSetup.LeftMargin = 0;
                sheet.PageSetup.RightMargin = 0;
                sheet.PageSetup.TopMargin = 0;
                sheet.PageSetup.BottomMargin = 0;

                // Reset fit‑to‑page options
                sheet.PageSetup.FitToPagesWide = 1;   // fit to one page wide
                sheet.PageSetup.FitToPagesTall = 0;   // unlimited pages tall

                // Reset other page‑setup flags to defaults
                sheet.PageSetup.CenterHorizontally = false;
                sheet.PageSetup.CenterVertically = false;
                sheet.PageSetup.PrintGridlines = false;
                sheet.PageSetup.PrintHeadings = false;
                // Orientation line removed because the enum may differ across versions
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the cleaned workbook as a new XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook cleaned and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
