// Title: Set 0.5‑inch margins and fit a worksheet to a single printed page with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set 0.5 inch left and right margins before printing | How to configure PageSetup to fit an Excel worksheet on one page using Aspose.Cells | Aspose.Cells .NET example for custom page margins and fit‑to‑page scaling | Apply uniform margins and fit-to-page in Aspose.Cells workbook C# code | Print layout settings for Excel sheet with 0.5 inch margins in Aspose.Cells
// Tags: Aspose.Cells set page margins C# | fit worksheet to one page Aspose.Cells | PageSetup margin configuration .NET | scale worksheet to fit printed page Aspose.Cells | custom margins inches Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, accesses the first worksheet, sets all margins (left, right, top, bottom) to 0.5 inches via the PageSetup object, configures FitToPagesWide and FitToPagesTall to 1 to scale the sheet to a single printed page, and saves the file as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the PageSetup object to configure printing options
                PageSetup pageSetup = worksheet.PageSetup;

                // Set custom margins of 0.5 inches on all sides
                pageSetup.LeftMargin = 0.5;   // inches
                pageSetup.RightMargin = 0.5;  // inches
                pageSetup.TopMargin = 0.5;    // inches
                pageSetup.BottomMargin = 0.5; // inches

                // Scale the worksheet to fit one page (both width and height)
                pageSetup.FitToPagesWide = 1;
                pageSetup.FitToPagesTall = 1;

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
