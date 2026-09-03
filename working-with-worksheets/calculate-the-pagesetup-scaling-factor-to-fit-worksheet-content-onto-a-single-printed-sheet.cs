// Title: Retrieve the page‑setup Zoom (scaling factor) after fitting an Excel worksheet to one printed page using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook, sets PageSetup.FitToPagesWide = 1 and FitToPagesTall = 1 with Aspose.Cells, then reads the resulting PageSetup.Zoom value and writes it to the console. | Show how to obtain the automatic scaling percentage after configuring a worksheet to fit on a single printed page in Aspose.Cells for .NET. | Provide a .NET snippet that saves the workbook after adjusting the page‑setup and logs the calculated scaling factor.
// Common Searches: Aspose.Cells .NET get page zoom after setting FitToPagesWide and FitToPagesTall | C# how to read scaling factor for single-page print in Excel using Aspose.Cells | retrieve worksheet print scaling percentage with Aspose.Cells for .NET | fit worksheet to one page and obtain Zoom value programmatically Aspose.Cells | Aspose.Cells page setup fit to one page calculate scaling factor
// Tags: Aspose.Cells page setup fit to one page | Aspose.Cells retrieve Zoom percentage | C# Excel worksheet print scaling Aspose.Cells | Aspose.Cells FitToPagesWide FitToPagesTall usage | Aspose.Cells calculate page scaling factor

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, configures its PageSetup to fit the content to a single printed page (both horizontally and vertically), reads the automatically computed Zoom (scaling) percentage, outputs the value, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Configure page setup to fit the entire content on a single printed sheet
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.FitToPagesWide = 1; // fit to 1 page horizontally
                pageSetup.FitToPagesTall = 1; // fit to 1 page vertically

                // After setting FitToPages, Aspose.Cells computes the scaling factor
                int scalingFactor = pageSetup.Zoom; // scaling percentage (e.g., 85 means 85%)

                // Output the scaling factor
                Console.WriteLine($"Scaling factor to fit on one page: {scalingFactor}%");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated page‑setup changes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
