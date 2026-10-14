// Title: Read worksheet paper size and compute custom print zoom to fit a target page using Aspose.Cells for .NET
// AI Prompts: Write C# code that extracts PageSetup.PaperWidth and PaperHeight, converts the values from points to inches, and calculates a zoom percentage to fit a specified target page size. | Create a reusable C# method that accepts target width and height in inches, computes the appropriate Worksheet.PageSetup.Zoom, and clamps the result between 10% and 400%. | Show how to output the original paper dimensions, scaling factors, and the applied zoom before saving the workbook.
// Common Searches: Aspose.Cells how to get worksheet paper width in inches | C# calculate print zoom from PageSetup.PaperHeight Aspose.Cells | Set worksheet zoom to match letter size using Aspose.Cells .NET API | Convert PageSetup paper dimensions from points to inches Aspose.Cells example | Clamp Aspose.Cells PageSetup.Zoom between 10 and 400 percent
// Tags: Aspose.Cells PageSetup paper size retrieval | worksheet print zoom calculation C# | convert points to inches Aspose.Cells | custom scaling factor for worksheet printing | set PageSetup.Zoom programmatically .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, reads the worksheet's PaperWidth and PaperHeight (in points), converts them to inches, computes X and Y scaling factors for a desired target page size, determines a zoom percentage (clamped to 10‑400%), applies it via PageSetup.Zoom, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Retrieve paper width and height in points (1 point = 1/72 inch)
            double paperWidthPoints = pageSetup.PaperWidth;
            double paperHeightPoints = pageSetup.PaperHeight;

            // Convert points to inches for easier calculations
            double paperWidthInches = paperWidthPoints / 72.0;
            double paperHeightInches = paperHeightPoints / 72.0;

            // Define desired custom dimensions (in inches)
            double targetWidthInches = 8.5;   // Example: Letter width
            double targetHeightInches = 11.0; // Example: Letter height

            // Calculate scaling factors based on target vs. actual paper size
            double scaleX = targetWidthInches / paperWidthInches;
            double scaleY = targetHeightInches / paperHeightInches;

            // Determine zoom percentage (must be between 10 and 400)
            int zoom = (int)(Math.Min(scaleX, scaleY) * 100);
            zoom = Math.Max(10, Math.Min(zoom, 400)); // Clamp to valid range

            // Apply the zoom
            pageSetup.Zoom = zoom;

            // Output the retrieved dimensions and calculated scaling factors
            Console.WriteLine($"Paper Width: {paperWidthInches:F2} inches");
            Console.WriteLine($"Paper Height: {paperHeightInches:F2} inches");
            Console.WriteLine($"Scale X: {scaleX:F3}");
            Console.WriteLine($"Scale Y: {scaleY:F3}");
            Console.WriteLine($"Applied Zoom: {pageSetup.Zoom}%");

            // Save the workbook if any changes need to be persisted
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
