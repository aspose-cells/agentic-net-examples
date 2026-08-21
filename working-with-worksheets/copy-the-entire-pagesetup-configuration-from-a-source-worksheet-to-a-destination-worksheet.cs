// Title: Copy worksheet page‑setup using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to configure page‑setup settings on a source worksheet and duplicate the entire configuration to another worksheet with the PageSetup.Copy method and default CopyOptions, then save the workbook.
// Keywords: Aspose.Cells page setup copy | PageSetup.Copy C# | duplicate worksheet print settings | Aspose.Cells copy options | transfer page orientation Aspose.Cells | C# workbook page layout
// Common Searches: Aspose.Cells copy page setup from one sheet to another | PageSetup.Copy example C# | How to duplicate print area in Aspose.Cells | Copy worksheet page layout Aspose.Cells .NET | Transfer page orientation between worksheets C#
// Developer Intent: Replicate all page‑setup properties from a source worksheet to a destination worksheet in a .NET workbook.
// Use Cases: Apply a predefined print layout to newly created worksheets in a report workbook. | Synchronize page‑setup settings before exporting multiple sheets to PDF or XPS. | Clone page‑setup when programmatically generating templated worksheets for automated reporting.
// AI Prompts: Generate C# code that copies page‑setup settings from one worksheet to many worksheets using a loop with Aspose.Cells. | Explain how to use custom CopyOptions with PageSetup.Copy to copy only selected properties. | Provide robust error handling for PageSetup.Copy when source or destination worksheets may be missing.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCopyDemo
{
    // Demonstrates how to configure page‑setup settings on a source worksheet and duplicate the entire configuration to another worksheet with the PageSetup.Copy method and default CopyOptions, then save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (source and destination are in the same workbook)
            Workbook workbook = new Workbook();

            // Access the first worksheet as the source
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";

            // Configure some page‑setup settings on the source worksheet
            PageSetup srcSetup = sourceSheet.PageSetup;
            srcSetup.PaperSize = PaperSizeType.PaperA3;
            srcSetup.Orientation = PageOrientationType.Landscape;
            srcSetup.PrintArea = "A1:D20";
            srcSetup.CenterHorizontally = true;
            srcSetup.CenterVertically = true;
            srcSetup.FitToPagesWide = 1;
            srcSetup.FitToPagesTall = 1;

            // Add a second worksheet that will receive the copied page‑setup
            Worksheet destSheet = workbook.Worksheets.Add("DestinationSheet");

            // Copy the entire page‑setup configuration from source to destination
            // Use the PageSetup.Copy method with default CopyOptions
            destSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Verify the copy (optional)
            Console.WriteLine("Destination paper size: " + destSheet.PageSetup.PaperSize);
            Console.WriteLine("Destination orientation: " + destSheet.PageSetup.Orientation);
            Console.WriteLine("Destination print area: " + destSheet.PageSetup.PrintArea);

            // Save the workbook (uses the provided save rule)
            workbook.Save("PageSetupCopyResult.xlsx");
        }
    }
}
