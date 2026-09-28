// Title: Apply black‑and‑white page setup to confidential worksheets before exporting to PDF with Aspose.Cells for .NET
// AI Prompts: Set the Worksheet.PageSetup.BlackAndWhite property to true for any sheet whose name contains 'Confidential' and then save the workbook as a PDF. | Modify the worksheet iteration to detect a naming pattern and enable grayscale printing only for those sheets before calling Workbook.Save with SaveFormat.Pdf.
// Common Searches: Aspose.Cells set black and white mode for specific worksheets before PDF export | apply black‑and‑white page setup to confidential Excel tabs using .NET | conditional page setup per worksheet when saving workbook as PDF with Aspose.Cells | export selected worksheets in black‑and‑white with Aspose.Cells .NET | how to filter worksheets by name for PDF generation in Aspose.Cells
// Tags: black‑and‑white page setup per worksheet | conditional worksheet PDF export Aspose.Cells | confidential sheet filtering .NET | apply printing options to selected worksheets | page setup property BlackAndWhite usage

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates through each worksheet, checks if the sheet name contains "Confidential", enables the BlackAndWhite page‑setup flag for those sheets, and then saves the entire workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet is marked as confidential (e.g., name contains "Confidential")
            if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Apply black‑and‑white mode for printing this sheet
                sheet.PageSetup.BlackAndWhite = true;
            }
        }

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
