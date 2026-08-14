// Title: Get Worksheet Paper Width & Height and Compute Custom Zoom in Aspose.Cells for .NET
// Description: Demonstrates how to read PageSetup.PaperWidth and PageSetup.PaperHeight (in inches) from a worksheet, calculate X/Y scaling ratios for a target page size, choose the smaller ratio to preserve aspect ratio, set the Zoom property, and save the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells get paper size | PageSetup PaperWidth C# | PageSetup PaperHeight C# | calculate worksheet scaling factor | custom zoom Aspose.Cells | aspect ratio scaling Excel | C# Aspose.Cells example | Excel print scaling programmatically
// Common Searches: how to retrieve paper width and height with Aspose.Cells | calculate custom zoom percentage from target page size | set worksheet zoom based on aspect ratio Aspose.Cells | PageSetup.PaperWidth property example | Aspose.Cells scaling for PDF export
// Developer Intent: Read the current paper dimensions of a worksheet and compute a zoom level that fits a specified target size while maintaining aspect ratio.
// Use Cases: Fit worksheet content to a predefined paper size (e.g., 8×10 in) before printing or PDF conversion. | Programmatically adjust zoom for dynamic reports where page layout must adapt to varying target dimensions. | Ensure consistent visual scaling across multiple workbooks generated in an automated reporting pipeline.
// AI Prompts: Show C# code that reads PageSetup.PaperWidth and PaperHeight, calculates X and Y scaling factors for an 8×10‑inch target, and applies the smaller factor as a percentage zoom in Aspose.Cells. | Explain how to handle division‑by‑zero when PaperWidth or PaperHeight is zero while computing scaling ratios. | Provide a step‑by‑step guide to preserve aspect ratio when setting the Zoom property based on custom page dimensions.

using System;
using Aspose.Cells;

// Demonstrates how to read PageSetup.PaperWidth and PageSetup.PaperHeight (in inches) from a worksheet, calculate X/Y scaling ratios for a target page size, choose the smaller ratio to preserve aspect ratio, set the Zoom property, and save the workbook using Aspose.Cells for C#.
class RetrievePaperSize
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the PageSetup object
        PageSetup pageSetup = sheet.PageSetup;

        // Retrieve paper width and height in inches
        double paperWidth = pageSetup.PaperWidth;
        double paperHeight = pageSetup.PaperHeight;

        // Display the retrieved dimensions
        Console.WriteLine($"Paper Width (inches): {paperWidth}");
        Console.WriteLine($"Paper Height (inches): {paperHeight}");

        // Example: calculate custom scaling factors for a target size
        double targetWidth = 8.0;   // desired width in inches
        double targetHeight = 10.0; // desired height in inches

        double scaleX = targetWidth / paperWidth;
        double scaleY = targetHeight / paperHeight;

        Console.WriteLine($"Scale X: {scaleX:F2}");
        Console.WriteLine($"Scale Y: {scaleY:F2}");

        // Apply a zoom based on the smaller scale to maintain aspect ratio
        double zoomPercent = Math.Min(scaleX, scaleY) * 100;
        pageSetup.Zoom = (int)Math.Round(zoomPercent);
        pageSetup.IsPercentScale = true;

        Console.WriteLine($"Applied Zoom (%): {pageSetup.Zoom}");

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
