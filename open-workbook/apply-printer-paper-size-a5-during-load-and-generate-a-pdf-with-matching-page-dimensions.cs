// Title: Set A5 printer paper size for every worksheet and export the workbook as an A5‑sized PDF using Aspose.Cells for .NET
// AI Prompts: Apply PaperSizeType.PaperA5 to each worksheet's PageSetup, enable FitToPagesWide = 1, then save the workbook as PDF. | Iterate through workbook.Worksheets, set pageSetup.PaperSize to A5, configure fit‑to‑width, and generate output.pdf with matching dimensions.
// Common Searches: How to set A5 page size for all sheets in Aspose.Cells before converting to PDF | Aspose.Cells .NET export Excel to PDF with A5 page dimensions | Fit worksheet content to page width when saving as A5 PDF using Aspose.Cells
// Tags: set worksheet paper size a5 Aspose.Cells | export workbook to pdf a5 page size | fit worksheet to page width Aspose.Cells | page setup paper size conversion .net

using Aspose.Cells;
using System;

// The example loads an existing Excel file, loops through each worksheet to set PageSetup.PaperSize to A5, configures the sheet to fit to one page wide with unlimited height, and then saves the workbook as a PDF whose pages match the A5 dimensions.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply A5 paper size to each worksheet's page setup
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PaperSize = PaperSizeType.PaperA5;   // Set printer paper size to A5
            // Optional: fit the content to the width of the page
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 0; // unlimited height
        }

        // Save the workbook as a PDF with the A5 page dimensions
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
