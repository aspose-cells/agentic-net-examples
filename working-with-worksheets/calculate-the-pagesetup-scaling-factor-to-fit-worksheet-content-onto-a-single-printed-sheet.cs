// Title: How to obtain the page‑setup scaling factor after fitting an Excel worksheet to a single printed page with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get page setup zoom after setting FitToPagesWide and FitToPagesTall | retrieve scaling percentage for one‑page print using Aspose.Cells PageSetup | determine print scaling factor programmatically with Aspose.Cells .NET | C# read calculated zoom value when fitting an Excel sheet to a single page in Aspose.Cells | how to fetch PageSetup.Zoom after applying FitToPages in Aspose.Cells
// Tags: Aspose.Cells PageSetup FitToPages scaling | retrieve PageSetup.Zoom value C# | fit worksheet to single printed page Aspose.Cells | calculate print scaling factor .NET | Excel workbook page setup zoom Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing workbook, accesses the first worksheet, sets PageSetup.FitToPagesWide and FitToPagesTall to 1 to force the content onto one printed page, reads the resulting PageSetup.Zoom property as the scaling factor (percentage), outputs it, and saves the workbook preserving the page‑setup configuration.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the page‑setup object
        PageSetup pageSetup = sheet.PageSetup;

        // Configure the page setup to fit the entire sheet on a single printed page
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 1;

        // After setting FitToPages, the Zoom property reflects the scaling factor (percentage)
        int scalingFactor = pageSetup.Zoom;

        // Output the scaling factor
        Console.WriteLine($"Scaling factor to fit on one printed sheet: {scalingFactor}%");

        // Save the workbook (optional, retains the page‑setup settings)
        workbook.Save("output.xlsx");
    }
}
