// Title: Clone Worksheet PageSetup and Adjust Paper Height with Aspose.Cells for .NET
// Description: Demonstrates how to copy the full PageSetup from one worksheet to another using PageSetup.Copy, then modify only the paper height via CustomPaperSize while preserving the original width, and finally save the workbook.
// Keywords: Aspose.Cells PageSetup copy | C# clone worksheet layout | custom paper size Aspose.Cells | PageSetup.Copy example | .NET print settings worksheet | modify paper height programmatically | Aspose.Cells tutorial | Aspose.Cells .NET US | Aspose.Cells Europe
// Common Searches: copy page setup from one sheet to another Aspose.Cells | change only paper height after cloning page setup | set custom paper dimensions C# Aspose.Cells | how to use PageSetup.Copy in .NET | adjust worksheet print size Aspose.Cells
// Developer Intent: Duplicate the page layout of a source worksheet on a target sheet and then change just the paper height.
// Use Cases: Reuse a template worksheet’s margins, orientation, and scaling across multiple reports while tailoring each sheet’s paper height for different print formats. | Generate a batch of printable sheets that share identical layout settings but require distinct paper heights such as Letter, Legal, or custom sizes. | Create automated Excel exports where the overall page setup is standardized, and only the height dimension varies per document.
// AI Prompts: Show C# code to copy a worksheet’s PageSetup to another sheet with Aspose.Cells and then set a custom paper height. | Explain how to use PageSetup.Copy and CustomPaperSize to preserve all layout settings except PaperHeight. | Provide a step‑by‑step example for cloning page setup and adjusting only the height in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCloneDemo
{
    // Demonstrates how to copy the full PageSetup from one worksheet to another using PageSetup.Copy, then modify only the paper height via CustomPaperSize while preserving the original width, and finally save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a second worksheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add(); // adds worksheet at index 1

            // Source worksheet (index 0) – set some initial page setup values
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA4;          // default A4
            sourceSheet.PageSetup.Orientation = PageOrientationType.Portrait;
            sourceSheet.PageSetup.FitToPagesWide = 1;
            sourceSheet.PageSetup.FitToPagesTall = 1;

            // Target worksheet (index 1) – will receive the cloned page setup
            Worksheet targetSheet = workbook.Worksheets[1];

            // Clone the page setup from source to target using Copy method
            // CopyOptions with default settings copies all properties
            targetSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Modify only the paper height of the target sheet.
            // PaperHeight is read‑only, so we use CustomPaperSize to set a new height
            // while preserving the existing width.
            double currentWidth = targetSheet.PageSetup.PaperWidth; // width in inches
            double newHeight = 11.0; // desired height in inches (e.g., Letter height)

            // Apply the custom paper size (width stays the same, height changes)
            targetSheet.PageSetup.CustomPaperSize(currentWidth, newHeight);

            // Save the workbook to verify the changes
            workbook.Save("ClonedPageSetup.xlsx");
        }
    }
}
