// Title: Export an Excel worksheet to a landscape-oriented PDF that fits to one page wide using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that configures a worksheet’s PageSetup to landscape, sets FitToPagesWide = 1 and FitToPagesTall = 0, and saves the workbook as a PDF with Aspose.Cells. | Write a C# example that prints an Excel sheet to a single-page-wide PDF in landscape mode using Aspose.Cells PageSetup properties.
// Common Searches: Aspose.Cells C# set worksheet orientation to landscape and fit to one page wide for PDF export | how to use FitToPagesWide and FitToPagesTall in Aspose.Cells when saving as PDF | export Excel to PDF landscape mode with single-page width using Aspose.Cells .NET | C# Aspose.Cells page setup for printable PDF fitting content to one page wide
// Tags: landscape orientation page setup Aspose.Cells | FitToPagesWide property PDF export C# | FitToPagesTall automatic height Aspose.Cells | export worksheet to PDF Aspose.Cells | page setup printable PDF .NET

using Aspose.Cells;
using System;

// // This program creates a workbook, sets the first worksheet's orientation to landscape, configures the page setup to fit the content to one page wide (height adjusts automatically), and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the worksheet to 1 page wide (height will adjust automatically)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0; // 0 means automatic

        // Save the workbook as a printable PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
