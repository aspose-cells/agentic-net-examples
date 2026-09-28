// Title: Set all worksheets to landscape, fit each to one page wide, and export the workbook as XPS using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, changes every worksheet's PageSetup to landscape orientation, sets FitToPagesWide to 1, and saves the result as an XPS document. | Show how to iterate through a Workbook's worksheets, apply page orientation and fit‑to‑page settings, and export the workbook to XPS format with Aspose.Cells.
// Common Searches: Aspose.Cells C# set page orientation landscape for all sheets and fit to one page width before saving as XPS | How to export an Excel workbook to XPS with landscape layout using Aspose.Cells .NET | Fit worksheet to single page width programmatically with Aspose.Cells and save as XPS
// Tags: Aspose.Cells set worksheet orientation landscape | FitToPagesWide page setup Aspose.Cells | export workbook to XPS Aspose.Cells | apply page settings to all worksheets .NET | Aspose.Cells SaveFormat.Xps usage

using Aspose.Cells;
using System;

// The example loads 'input.xlsx', iterates through each worksheet to set the PageSetup orientation to landscape, configures FitToPagesWide = 1 (FitToPagesTall = 0) so the content fits one page wide, and then saves the workbook as 'output.xps' using the XPS save format.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure page settings for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            // Set orientation to landscape
            pageSetup.Orientation = PageOrientationType.Landscape;
            // Fit the sheet to one page wide; height will adjust automatically
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 0;
        }

        // Save the workbook as XPS
        workbook.Save("output.xps", SaveFormat.Xps);
    }
}
