// Title: Set Worksheet Page Zoom to 80% and Export as PDF with Aspose.Cells for .NET
// Description: Demonstrates how to set PageSetup.Zoom to 80 % (with IsPercentScale enabled) on a worksheet and save the workbook directly to a PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | PageSetup.Zoom | IsPercentScale | PDF export | worksheet zoom | set page zoom 80 percent | save workbook as PDF
// Common Searches: Aspose.Cells set worksheet zoom before PDF export | C# PageSetup.Zoom 80 percent Aspose.Cells | How to enable percent scaling in Aspose.Cells | Export workbook to PDF with custom zoom Aspose | Aspose.Cells PDF scaling options
// Developer Intent: Apply an 80 % page zoom to a worksheet and generate a PDF document.
// Use Cases: Create a new workbook, configure PageSetup.Zoom = 80 and PageSetup.IsPercentScale = true, then call Workbook.Save with SaveFormat.Pdf to produce a scaled PDF. | Add sample data to verify that the 80 % zoom is reflected in the exported PDF. | Batch process multiple worksheets, each with a different zoom level, before combining them into a single PDF.
// AI Prompts: Write C# code that sets a worksheet's page zoom to a specific percentage and saves the workbook as a PDF using Aspose.Cells. | Show how to enable percent scaling (IsPercentScale) and adjust Zoom for PDF output in Aspose.Cells for .NET. | Explain how to apply different zoom levels to several worksheets and export them as one PDF file with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set PageSetup.Zoom to 80 % (with IsPercentScale enabled) on a worksheet and save the workbook directly to a PDF file using Aspose.Cells for .NET.
class SetZoomAndExportPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the page setup zoom to 80%
        worksheet.PageSetup.Zoom = 80;
        // Ensure the zoom is interpreted as a percent scale
        worksheet.PageSetup.IsPercentScale = true;

        // (Optional) Add some sample data to visualize the scaling
        worksheet.Cells["A1"].PutValue("Worksheet with 80% zoom");

        // Export the workbook to PDF
        workbook.Save("Zoom80.pdf", SaveFormat.Pdf);
    }
}
