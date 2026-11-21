// Title: Apply black‑and‑white page setup to confidential worksheets before saving as PDF using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set BlackAndWhite page setup for worksheets that contain 'Confidential' in the name | How to export only selected Excel sheets as black‑and‑white PDF with Aspose.Cells .NET | Conditional worksheet formatting before PDF conversion in Aspose.Cells C# | Apply black and white printing mode to specific worksheets when generating PDF with Aspose.Cells | Filter worksheets by name and enable black‑and‑white mode for PDF export in C#
// Tags: conditional black and white page setup Aspose.Cells | confidential worksheet PDF export Aspose.Cells | worksheet name filter Aspose.Cells | black and white PDF conversion Aspose.Cells | selective worksheet formatting C#

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates through each worksheet, enables the BlackAndWhite page‑setup flag for sheets whose names contain "Confidential", and then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Identify confidential sheets (e.g., by name containing "Confidential")
            if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Apply black‑and‑white mode for printing/PDF export
                sheet.PageSetup.BlackAndWhite = true;
            }
        }

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
