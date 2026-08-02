// Title: Get Worksheet Paper Size and Apply Custom Scaling with Aspose.Cells for .NET (C#)
// Description: Read PaperWidth and PaperHeight from a worksheet's PageSetup, compute width/height ratios for a target printable area, choose the smaller ratio as a custom zoom percentage, enable IsPercentScale, and save the workbook.
// Keywords: Aspose.Cells | C# | PageSetup | PaperWidth | PaperHeight | custom scaling | worksheet zoom | printable area | Excel to PDF scaling | IsPercentScale
// Common Searches: Aspose.Cells get paper size | calculate worksheet zoom programmatically | set custom scale based on printable area | PageSetup PaperWidth C# | adjust Excel print scaling Aspose
// Developer Intent: Read the worksheet's paper dimensions and compute a percentage zoom that fits a specified printable region.
// Use Cases: Fit a 6×8‑inch region onto any paper size before printing or exporting. | Generate PDFs with consistent content scaling across different page formats. | Dynamically adapt worksheet zoom for custom report layouts. | Ensure content stays within page margins when exporting to images.
// AI Prompts: Provide C# code using Aspose.Cells to read PaperWidth and PaperHeight from a worksheet's PageSetup and calculate a zoom factor for a 6"×8" printable area. | Show how to apply the smaller scaling ratio as a percentage Zoom and enable IsPercentScale in Aspose.Cells. | Explain how to handle cases where the desired printable size exceeds the actual paper dimensions and suggest fallback strategies.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Read PaperWidth and PaperHeight from a worksheet's PageSetup, compute width/height ratios for a target printable area, choose the smaller ratio as a custom zoom percentage, enable IsPercentScale, and save the workbook.
    class RetrievePaperSizeAndCalculateScale
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the PageSetup object of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Retrieve the paper width and height (in inches) from PageSetup
            double paperWidthInches = pageSetup.PaperWidth;   // property: PaperWidth
            double paperHeightInches = pageSetup.PaperHeight; // property: PaperHeight

            // Example: Desired printable area size (in inches)
            double desiredWidthInches = 6.0;
            double desiredHeightInches = 8.0;

            // Calculate scaling factors based on desired size vs actual paper size
            double widthScale = desiredWidthInches / paperWidthInches;
            double heightScale = desiredHeightInches / paperHeightInches;

            // Choose the smaller scale to fit both dimensions
            double customScale = Math.Min(widthScale, heightScale);

            // Output the retrieved sizes and calculated scale
            Console.WriteLine($"Paper Width (inches): {paperWidthInches}");
            Console.WriteLine($"Paper Height (inches): {paperHeightInches}");
            Console.WriteLine($"Desired Width (inches): {desiredWidthInches}");
            Console.WriteLine($"Desired Height (inches): {desiredHeightInches}");
            Console.WriteLine($"Width Scale: {widthScale:P2}");
            Console.WriteLine($"Height Scale: {heightScale:P2}");
            Console.WriteLine($"Custom Scale (minimum): {customScale:P2}");

            // Apply the custom scale to the worksheet (as a percentage)
            pageSetup.Zoom = (int)(customScale * 100);
            pageSetup.IsPercentScale = true;

            // Save the workbook (demonstrates usage of save lifecycle)
            workbook.Save("ScaledWorkbook.xlsx");
        }
    }
}
