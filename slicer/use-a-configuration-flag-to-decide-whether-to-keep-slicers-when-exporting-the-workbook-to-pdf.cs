// Title: C# – Export Excel to PDF with optional slicer visibility using Aspose.Cells
// Description: Loads a workbook, toggles each slicer's IsPrintable property based on a Boolean flag, and saves the file as PDF with PdfSaveOptions. The flag lets developers include or hide slicers in the generated PDF.
// Keywords: Aspose.Cells | C# PDF export | Excel slicer visibility | IsPrintable property | toggle slicers | conditional PDF rendering | PdfSaveOptions | export workbook to PDF | hide slicers Aspose | include slicers PDF
// Common Searches: Aspose.Cells hide slicers when saving to PDF | C# set slicer IsPrintable before PDF export | how to exclude slicers from PDF using Aspose | toggle slicer visibility in PDF output | conditional slicer export Aspose.Cells
// Developer Intent: Provide a simple way to control whether slicers appear in a PDF generated from an Excel workbook by using a configuration flag.
// Use Cases: Create client‑ready PDFs without interactive slicers for a clean look. | Produce internal documentation that retains slicers to demonstrate filtering options. | Offer end‑users a setting to export reports with or without slicers based on privacy or presentation preferences.
// AI Prompts: Generate a C# method that accepts a bool parameter to include slicers and exports an Excel workbook to PDF with Aspose.Cells. | Show code that iterates through all worksheets, sets each slicer's IsPrintable property according to a flag, and saves the workbook as PDF. | Explain how to combine PdfSaveOptions with other rendering settings while respecting a slicer‑visibility flag.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

// Loads a workbook, toggles each slicer's IsPrintable property based on a Boolean flag, and saves the file as PDF with PdfSaveOptions. The flag lets developers include or hide slicers in the generated PDF.
class ExportWorkbookToPdfWithSlicerOption
{
    static void Main()
    {
        // Configuration flag: true to keep slicers in the PDF, false to hide them
        bool keepSlicers = true; // Change as needed

        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and their slicers
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet contains any slicers
            if (sheet.Slicers.Count > 0)
            {
                for (int i = 0; i < sheet.Slicers.Count; i++)
                {
                    Slicer slicer = sheet.Slicers[i];
                    // The IsPrintable property determines whether the slicer appears in the PDF
                    slicer.IsPrintable = keepSlicers;
                }
            }
        }

        // Create PDF save options (additional options can be set here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
