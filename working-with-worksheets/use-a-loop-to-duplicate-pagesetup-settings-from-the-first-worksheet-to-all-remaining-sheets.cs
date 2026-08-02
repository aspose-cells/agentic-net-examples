// Title: Copy PageSetup Settings from the First Worksheet to All Sheets with Aspose.Cells for .NET
// Description: This example creates a workbook with three worksheets, configures page‑setup properties (A3 paper, landscape orientation, fit‑to‑width, print title rows/columns) on the first sheet, then loops through the remaining sheets and copies those settings using PageSetup.Copy before saving the file.
// Keywords: Aspose.Cells page setup copy C# | duplicate worksheet print settings .NET | PageSetup.Copy example | apply same page layout to multiple sheets | loop copy page setup Aspose.Cells
// Common Searches: copy page setup from one worksheet to others Aspose.Cells | Aspose.Cells loop duplicate print settings | C# copy worksheet page layout | apply same paper size to all sheets programmatically
// Developer Intent: Replicate the page‑setup configuration of the first worksheet across every other worksheet in the workbook.
// Use Cases: Set A3 landscape and print titles on a template sheet, then propagate the layout to all report sheets for consistent printing. | Generate a multi‑sheet workbook where each sheet must share identical fit‑to‑page and orientation settings before exporting to PDF. | Automate uniform page‑setup across worksheets when producing batch reports that require the same print format.
// AI Prompts: Write C# code using Aspose.Cells that copies the page‑setup settings from the first worksheet to all other worksheets in a loop. | Explain which page‑setup properties are transferred by PageSetup.Copy when default CopyOptions are used. | Show an example that copies page‑setup settings to every sheet and then changes the orientation of a specific sheet afterwards.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCopyDemo
{
    // This example creates a workbook with three worksheets, configures page‑setup properties (A3 paper, landscape orientation, fit‑to‑width, print title rows/columns) on the first sheet, then loops through the remaining sheets and copies those settings using PageSetup.Copy before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Configure page‑setup settings on the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            PageSetup firstSetup = firstSheet.PageSetup;
            firstSetup.PaperSize = PaperSizeType.PaperA3;
            firstSetup.Orientation = PageOrientationType.Landscape;
            firstSetup.FitToPagesWide = 1;
            firstSetup.FitToPagesTall = 0;
            firstSetup.PrintTitleRows = "$1:$1";
            firstSetup.PrintTitleColumns = "$A:$A";

            // Loop through all remaining worksheets and copy the page‑setup settings
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                Worksheet targetSheet = workbook.Worksheets[i];
                // Use PageSetup.Copy to duplicate settings from the first sheet
                targetSheet.PageSetup.Copy(firstSetup, new CopyOptions());
            }

            // Save the workbook to a file
            workbook.Save("PageSetupCopied.xlsx");
        }
    }
}
