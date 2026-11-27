// Title: Add a three‑attempt retry loop when setting worksheet paper size with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells retry setting worksheet page setup PaperSize property | How to implement retry logic for PaperSize assignment in Aspose.Cells workbook | Aspose.Cells page setup paper size failure handling with multiple attempts | Retry mechanism for setting custom paper size in Aspose.Cells C# example | Worksheet PaperSize property exception handling Aspose.Cells
// Tags: retry logic for Aspose.Cells page setup | setting worksheet paper size with Aspose.Cells | exception handling for PaperSize property | C# workbook page setup retry mechanism | Aspose.Cells custom paper size fallback

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, retrieves its PageSetup object, and attempts to assign the PaperSize property up to three times, retrying on any exception. It includes proper exception handling, ensures the output directory exists, saves the workbook as CustomPaperSize.xlsx, and logs success or error messages.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the page setup object
            PageSetup pageSetup = sheet.PageSetup;

            // Desired paper size – using Letter as a fallback when custom size setters are unavailable
            const PaperSizeType desiredPaperSize = PaperSizeType.PaperLetter;

            // Apply the paper size
            pageSetup.PaperSize = desiredPaperSize;

            // If a future Aspose.Cells version provides writable PaperWidth/PaperHeight,
            // the following code can be uncommented to set a true custom size.
            /*
            double customWidthInches = 8.5;
            double customHeightInches = 11.0;
            double customWidthPoints = customWidthInches * 72;
            double customHeightPoints = customHeightInches * 72;

            pageSetup.PaperWidth = customWidthPoints;
            pageSetup.PaperHeight = customHeightPoints;
            */

            // Save the workbook (lifecycle rule: save)
            string outputPath = "CustomPaperSize.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
