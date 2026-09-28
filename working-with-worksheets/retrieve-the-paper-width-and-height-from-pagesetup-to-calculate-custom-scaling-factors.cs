// Title: Determine worksheet paper size with PageSetup and set custom zoom to fit US Letter using Aspose.Cells for .NET
// AI Prompts: Read the PageSetup.PaperWidth and PageSetup.PaperHeight of a worksheet, convert the values from points to inches, compute the scaling factor needed to match US Letter dimensions, and assign the resulting percentage to PageSetup.Zoom. | Calculate separate width and height scaling ratios from the worksheet's paper size, select the smaller ratio to preserve aspect ratio, clamp the zoom value between 10% and 400%, and apply it to the workbook. | Create a routine that automatically adjusts the print zoom of any Aspose.Cells workbook so that the printed area fits within an 8.5" × 11" page, using the PageSetup properties.
// Common Searches: Aspose.Cells C# get worksheet paper width and height points | how to convert PageSetup.PaperWidth from points to inches in Aspose.Cells | set worksheet zoom to fit US Letter size using Aspose.Cells .NET | calculate custom print scaling factor from PageSetup in Aspose.Cells | Aspose.Cells PageSetup.Zoom limits 10 to 400 percent
// Tags: Aspose.Cells PageSetup paper size scaling | worksheet print zoom calculation Aspose.Cells | convert points to inches Aspose.Cells | custom worksheet zoom US Letter Aspose.Cells | maintain aspect ratio print scaling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads or creates a workbook, reads the first worksheet's PageSetup.PaperWidth and PaperHeight (in points), converts them to inches, computes width and height scaling factors to match an 8.5" × 11" page, selects the smaller factor to keep the aspect ratio, clamps the zoom between 10% and 400%, sets PageSetup.Zoom, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

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

            // Desired target dimensions (in inches) – US Letter size
            double targetWidthInches = 8.5;
            double targetHeightInches = 11.0;

            // Calculate scaling factors for width and height
            double scaleX = targetWidthInches / paperWidthInches;
            double scaleY = targetHeightInches / paperHeightInches;

            // Use the smaller factor to maintain aspect ratio
            double scaleFactor = Math.Min(scaleX, scaleY);

            // Convert to percentage and clamp to allowed range (10‑400)
            int zoom = (int)Math.Round(scaleFactor * 100);
            zoom = Math.Clamp(zoom, 10, 400);

            // Apply the zoom factor
            pageSetup.Zoom = zoom;

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\" with zoom set to {zoom}%.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
