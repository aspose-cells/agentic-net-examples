// Title: Copy PageSetup Settings from the First Worksheet to All Sheets with Aspose.Cells for .NET
// Description: Demonstrates how to configure PageSetup on the first worksheet (paper size, orientation, fit‑to‑page, print titles) and then use a C# loop with PageSetup.Copy and CopyOptions to apply the same settings to every other worksheet in the workbook before saving.
// Keywords: Aspose.Cells PageSetup copy C# | duplicate worksheet print settings .NET | loop copy page setup Aspose | PageSetup.Copy example | Aspose.Cells printing layout automation
// Common Searches: Aspose.Cells copy page setup to other sheets | C# loop to duplicate worksheet print settings | How to apply same PageSetup to all worksheets in Aspose.Cells | PageSetup.Copy with CopyOptions in .NET
// Developer Intent: Apply the PageSetup configuration of the first worksheet to every remaining worksheet in a workbook programmatically.
// Use Cases: Standardize printing layout (paper size, orientation, fit‑to‑page, titles) across all sheets in multi‑sheet reports. | Automatically propagate page‑setup settings when new worksheets are added at runtime. | Maintain consistent print behavior without manually configuring each worksheet.
// AI Prompts: Generate C# code using Aspose.Cells that copies the first worksheet's PageSetup to all existing and future worksheets. | Show how to use PageSetup.Copy with custom CopyOptions to duplicate only selected properties such as print titles. | Explain how to modify the loop to exclude hidden worksheets while copying PageSetup settings.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCopyDemo
{
    // Demonstrates how to configure PageSetup on the first worksheet (paper size, orientation, fit‑to‑page, print titles) and then use a C# loop with PageSetup.Copy and CopyOptions to apply the same settings to every other worksheet in the workbook before saving.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook with default worksheets
            Workbook workbook = new Workbook();

            // Access the first worksheet and configure its page setup
            Worksheet firstSheet = workbook.Worksheets[0];
            PageSetup firstPageSetup = firstSheet.PageSetup;
            firstPageSetup.PaperSize = PaperSizeType.PaperA4;
            firstPageSetup.Orientation = PageOrientationType.Landscape;
            firstPageSetup.FitToPagesWide = 1;
            firstPageSetup.FitToPagesTall = 0; // Let height adjust automatically
            firstPageSetup.PrintTitleRows = "$1:$1";
            firstPageSetup.PrintTitleColumns = "$A:$A";

            // Ensure there are additional worksheets to demonstrate copying
            // (Add two more sheets for the example)
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Loop through all worksheets except the first one
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                // Copy the page‑setup settings from the first worksheet to the current worksheet
                workbook.Worksheets[i].PageSetup.Copy(firstPageSetup, new CopyOptions());
            }

            // Save the workbook to verify the result
            workbook.Save("PageSetupCopied.xlsx");
        }
    }
}
