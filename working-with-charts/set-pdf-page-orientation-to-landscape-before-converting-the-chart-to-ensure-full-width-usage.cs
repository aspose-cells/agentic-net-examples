// Title: Export an Excel chart to a full‑width PDF by setting the worksheet orientation to landscape with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets a worksheet's PageSetup orientation to Landscape and then saves the workbook as a PDF using Aspose.Cells. | Show how to configure PageSetup to fit a chart to one page wide before converting the sheet to PDF in Aspose.Cells. | Provide a step‑by‑step example of exporting a chart from an Excel file to a landscape PDF with full page width using Aspose.Cells for .NET.
// Common Searches: how to set landscape orientation for a worksheet before exporting to PDF with Aspose.Cells C# | Aspose.Cells C# fit chart to page width when saving as PDF | export Excel chart to PDF full width landscape Aspose.Cells example | C# code to change PageSetup orientation and fit chart to one page wide in Aspose.Cells
// Tags: worksheet pageorientation landscape Aspose.Cells | chart fit topagewide pdf Aspose.Cells | export chart to pdf landscape C# | page setup orientation pdf conversion Aspose.Cells | fit chart to one page wide Aspose.Cells

using Aspose.Cells;
using System;

// Loads an Excel workbook, sets the first worksheet's PageSetup orientation to Landscape, optionally configures FitToPagesWide to 1 for full‑width chart rendering, and saves the sheet as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the chart
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet where the chart resides (e.g., the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page orientation to Landscape before PDF conversion
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Optional: fit the chart to the page width for full‑width usage
        sheet.PageSetup.FitToPagesWide = 1;   // fit to one page wide
        sheet.PageSetup.FitToPagesTall = 0;   // unlimited height

        // Convert the worksheet (including the chart) to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
