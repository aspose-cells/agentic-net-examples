// Title: Configure Aspose.Cells worksheet to fit all columns on a single printed page without scaling distortion (C#)
// Common Searches: Aspose.Cells C# set worksheet to fit all columns on one printed page | How to prevent scaling distortion when printing Excel with Aspose.Cells .NET | FitToPagesWide = 1 and FitToPagesTall = 0 example using Aspose.Cells | Set worksheet page setup zoom to 100% in Aspose.Cells C# | Print Excel worksheet with columns fitting on one page using Aspose.Cells API
// Tags: Aspose.Cells worksheet page setup fit to pages wide | C# set FitToPagesTall zero Aspose.Cells | Aspose.Cells print scaling 100 percent | Excel column fit single page Aspose.Cells | Aspose.Cells page setup zoom property

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, accesses the first worksheet, and configures its PageSetup so that all columns fit on a single printed page by setting FitToPagesWide to 1 and FitToPagesTall to 0, while preserving original size with a Zoom of 100%. It ensures the output directory exists and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet sheet = workbook.Worksheets[0];

            // Access the page setup of the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Fit all columns on a single page (wide)
            pageSetup.FitToPagesWide = 1;

            // Allow any number of pages tall (0 means unlimited)
            pageSetup.FitToPagesTall = 0;

            // Set scaling to 100% (no distortion) using Zoom property
            pageSetup.Zoom = 100;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
