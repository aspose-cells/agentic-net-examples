// Title: Configure all worksheets to A4 paper size and landscape orientation before saving as PDF using Aspose.Cells for .NET
// AI Prompts: Set each worksheet's PageSetup.PaperSize to PaperA4 and PageSetup.Orientation to Landscape, then export the workbook to PDF with Aspose.Cells. | Iterate through a workbook's worksheets, apply A4 landscape page settings, and call Workbook.Save using SaveFormat.Pdf in C#. | Programmatically configure PDF page size to A4 and orientation to landscape for every sheet and generate the PDF file with Aspose.Cells.
// Common Searches: Aspose.Cells export workbook to PDF with A4 landscape page size | C# set PDF page orientation to landscape for all worksheets using Aspose.Cells | change paper size to A4 for PDF output in Aspose.Cells .NET | apply same PDF page setup to every sheet before saving with Aspose.Cells | Aspose.Cells PDF export page layout loop worksheets
// Tags: Aspose.Cells PDF A4 landscape page setup | set worksheet PageSetup paper size A4 | apply landscape orientation Aspose.Cells PDF export | C# workbook save as PDF with custom page settings | iterate worksheets configure PDF page layout Aspose.Cells

using System;
using Aspose.Cells;

// The example creates or loads a workbook, loops through all worksheets to set the PageSetup to A4 paper size and landscape orientation, and then saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Iterate through all worksheets to set PDF page settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set paper size to A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            // Set orientation to landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;
        }

        // Save the workbook as a PDF with the configured page settings
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
