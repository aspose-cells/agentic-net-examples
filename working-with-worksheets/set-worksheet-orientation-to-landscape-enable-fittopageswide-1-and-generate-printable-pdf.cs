// Title: Export a worksheet to PDF in landscape orientation with FitToPagesWide = 1 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets a worksheet’s PageSetup to landscape, configures FitToPagesWide = 1 and FitToPagesTall = 0, then saves the workbook as a PDF with Aspose.Cells. | Show how to adjust page setup for printable PDF output so the content fits one page wide while keeping height automatic in Aspose.Cells.
// Common Searches: Aspose.Cells C# set page orientation to landscape and fit worksheet to one page wide before PDF export | How to use FitToPagesWide and FitToPagesTall in Aspose.Cells to create a printable PDF | Export Excel worksheet to PDF with landscape layout and automatic height using Aspose.Cells .NET
// Tags: set worksheet landscape orientation Aspose.Cells | fit worksheet to one page wide PDF Aspose.Cells | configure FitToPagesTall automatic height Aspose.Cells | export workbook to PDF with page setup Aspose.Cells

using Aspose.Cells;
using System;

// Creates a workbook, sets the first worksheet to landscape orientation, configures the page setup to fit the content to one page wide with automatic height, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page orientation to landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the sheet to 1 page wide (height will adjust automatically)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0; // 0 means automatic

        // Save the workbook as a printable PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
