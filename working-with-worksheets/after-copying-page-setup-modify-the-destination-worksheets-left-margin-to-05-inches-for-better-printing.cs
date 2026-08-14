// Title: Copy worksheet page setup and set left margin to 0.5 inches with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook with two worksheets, configures the source sheet’s page setup (paper size, orientation, left margin), copies the entire page setup to a second sheet using Worksheet.PageSetup.Copy, then overrides the left margin to 0.5 inches before saving the file.
// Keywords: Aspose.Cells | C# | .NET | copy page setup | worksheet left margin | 0.5 inches | PageSetup.Copy | CopyOptions | Excel printing | modify margins programmatically
// Common Searches: Aspose.Cells copy page setup C# | set left margin 0.5 inches Aspose.Cells | change worksheet margin after copying page setup | copy page setup between worksheets .NET | adjust left margin for printing Aspose.Cells
// Developer Intent: Copy a worksheet’s page setup to another sheet and then change the destination sheet’s left margin to 0.5 inches.
// Use Cases: Generate multi‑sheet reports where all sheets share the same paper size and orientation, but specific sheets need a narrower left margin for binding. | Automate creation of printable Excel files that reuse a base page setup while customizing margins per sheet to fit more content. | Standardize page layout across worksheets in a workbook and then fine‑tune the left margin for sheets that require different printable widths.
// AI Prompts: Show C# code that copies a worksheet’s page setup with Aspose.Cells and sets the destination left margin to 0.5 inches. | Provide an Aspose.Cells .NET example that uses Worksheet.PageSetup.Copy and then updates only the LeftMarginInch property. | Explain how to preserve all page‑setup settings while overriding the left margin after copying between worksheets in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupExample
{
    // C# example that creates a workbook with two worksheets, configures the source sheet’s page setup (paper size, orientation, left margin), copies the entire page setup to a second sheet using Worksheet.PageSetup.Copy, then overrides the left margin to 0.5 inches before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add(); // Adds a second worksheet (index 1)

            // Configure page setup for the source worksheet (index 0)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            sourceSheet.PageSetup.Orientation = PageOrientationType.Portrait;
            sourceSheet.PageSetup.LeftMarginInch = 1.0; // Example original left margin

            // Copy page setup from source worksheet to destination worksheet (index 1)
            Worksheet destSheet = workbook.Worksheets[1];
            destSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Modify the left margin of the destination worksheet to 0.5 inches
            destSheet.PageSetup.LeftMarginInch = 0.5;

            // Save the workbook to a file
            workbook.Save("PageSetupCopyAndMarginModified.xlsx");
        }
    }
}
