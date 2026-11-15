// Title: How to read worksheet paper width and height from PageSetup and apply custom zoom scaling with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get worksheet paper width and height from PageSetup | calculate zoom percentage to fit Excel sheet to Letter size using Aspose.Cells | set custom print scaling based on paper dimensions in Aspose.Cells .NET | convert PageSetup PaperWidth points to inches in C# | apply zoom limits 10-400% with Aspose.Cells PageSetup
// Tags: PageSetup.PaperWidth retrieval Aspose.Cells | custom zoom calculation Aspose.Cells .NET | points to inches conversion C# | worksheet print scaling Aspose.Cells | set PageSetup.Zoom Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one if missing), accesses the first worksheet's PageSetup, reads PaperWidth and PaperHeight in points, converts those values to inches, computes scaling factors to match a desired Letter size, selects the smaller factor, clamps the resulting zoom to the allowed 10‑400% range, assigns it to PageSetup.Zoom, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a new workbook as a fallback.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                Console.WriteLine($"Input file \"{inputPath}\" not found. A new workbook has been created.");
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object.
            PageSetup pageSetup = worksheet.PageSetup;

            // Retrieve paper dimensions in points (1 point = 1/72 inch).
            double paperWidthPoints = pageSetup.PaperWidth;
            double paperHeightPoints = pageSetup.PaperHeight;

            // Convert points to inches.
            double paperWidthInches = paperWidthPoints / 72.0;
            double paperHeightInches = paperHeightPoints / 72.0;

            // Desired dimensions (adjust as needed).
            double desiredWidthInches = 8.5;   // standard Letter width
            double desiredHeightInches = 11.0; // standard Letter height

            // Calculate scaling factors.
            double scaleX = desiredWidthInches / paperWidthInches;
            double scaleY = desiredHeightInches / paperHeightInches;

            // Choose the smaller factor to keep content within both dimensions.
            double scaleFactor = Math.Min(scaleX, scaleY);

            // Convert to zoom percentage and clamp to the allowed range (10‑400).
            int zoom = (int)Math.Round(scaleFactor * 100);
            zoom = Math.Max(10, Math.Min(400, zoom));

            // Apply the zoom.
            pageSetup.Zoom = zoom;

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\" with zoom set to {zoom}%.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
