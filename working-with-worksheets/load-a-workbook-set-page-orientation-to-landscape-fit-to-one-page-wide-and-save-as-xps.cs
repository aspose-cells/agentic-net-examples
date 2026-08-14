// Title: Aspose.Cells for .NET: Set Landscape Orientation, Fit to One Page Wide, Export to XPS
// Description: Loads an Excel workbook, changes every worksheet to landscape mode, configures the page setup to fit the content to one page width (height auto‑scaled), optionally forces one page per sheet, and saves the result as an XPS document using XpsSaveOptions.
// Keywords: Aspose.Cells | C# | .NET | landscape orientation | fit to one page wide | XPS export | page setup | OnePagePerSheet | Excel to XPS
// Common Searches: Aspose.Cells set worksheet orientation to landscape | fit Excel sheet to one page width before XPS export | save workbook as XPS with Aspose.Cells .NET | force one page per sheet XPS using Aspose.Cells | C# code to export Excel to XPS with page scaling
// Developer Intent: Apply landscape layout, fit each sheet to a single page width, and generate an XPS file from an Excel workbook.
// Use Cases: Create printable XPS reports that maintain a consistent landscape layout across multiple sheets. | Produce single‑page XPS files per worksheet for easy distribution or archiving. | Generate XPS documents that automatically scale to page width, eliminating manual print adjustments.
// AI Prompts: Generate C# code with Aspose.Cells to set all worksheets to portrait orientation, fit to one page tall, and export to PDF. | Show how to export a workbook to XPS with custom margins and without the OnePagePerSheet option using Aspose.Cells. | Explain the steps to adjust scaling, orientation, and page setup in Aspose.Cells before saving a workbook as XPS.

using System;
using Aspose.Cells;

// Loads an Excel workbook, changes every worksheet to landscape mode, configures the page setup to fit the content to one page width (height auto‑scaled), optionally forces one page per sheet, and saves the result as an XPS document using XpsSaveOptions.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure each worksheet's page setup
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set orientation to Landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Fit the worksheet to one page wide; height will adjust automatically
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;
        }

        // Create XPS save options (using the provided constructor rule)
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Optional: force each sheet to be rendered on a single page
        saveOptions.OnePagePerSheet = true;

        // Save the workbook as XPS with the configured options
        workbook.Save("output.xps", saveOptions);
    }
}
