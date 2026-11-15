// Title: Set worksheet orientation to landscape, fit each sheet to one page width, and export the workbook as XPS using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet orientation to landscape and fit to one page width | How to export an Excel workbook to XPS with Aspose.Cells .NET | Apply page setup settings to all worksheets in Aspose.Cells C# | Fit Excel sheet to a single page wide using Aspose.Cells API | Save workbook as XPS file programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells set landscape orientation | Aspose.Cells configure page scaling | Aspose.Cells save as XPS format | Aspose.Cells apply page setup to all worksheets | C# export Excel to XPS using Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing Excel file, iterates through every worksheet to set the page orientation to landscape and configure the page setup to fit the content to one page wide (height adjusts automatically), and then saves the workbook as an XPS document using the Aspose.Cells SaveFormat.Xps option.
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
            // Fit the sheet to one page wide (height will adjust automatically)
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 0;
        }

        // Save the workbook as XPS
        workbook.Save("output.xps", SaveFormat.Xps);
    }
}
