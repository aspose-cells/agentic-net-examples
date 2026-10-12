// Title: How to set a worksheet to landscape orientation, fit it to one page wide, and export the workbook as XPS using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, changes the first worksheet's page orientation to landscape, configures the page setup to fit the sheet to a single page width, and saves the workbook as an XPS document. | Show how to use Aspose.Cells PageSetup in a .NET application to print a worksheet in landscape mode, fit it to one page wide, and then export the workbook to XPS format.
// Common Searches: aspnet aspose.cells set landscape orientation and fit to one page wide before saving as xps | c# how to export excel to xps with page setup using Aspose.Cells | fit worksheet to single page width landscape Aspose.Cells .NET | save workbook as XPS file after adjusting page layout Aspose.Cells
// Tags: landscape orientation page setup Aspose.Cells | single-page-width printing Aspose.Cells | export workbook to XPS Aspose.Cells | configure worksheet page layout .NET | save Excel as XPS C#

using Aspose.Cells;
using System;

// // Loads 'input.xlsx', sets the first worksheet's orientation to landscape, configures the page setup to fit the sheet to one page wide (height automatic), and saves the workbook as 'output.xps' in XPS format.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can loop through all worksheets if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the sheet to one page wide (height will be automatic)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // Save the workbook as an XPS document
        workbook.Save("output.xps", SaveFormat.Xps);
    }
}
