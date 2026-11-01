// Title: Set black‑and‑white printing for worksheets named “Confidential” and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: In C#, iterate through a Workbook’s Worksheets, enable Worksheet.PageSetup.BlackAndWhite for each sheet whose Name contains the word “Confidential”, then save the workbook as a PDF. | Use Aspose.Cells to apply a conditional black‑and‑white page setup to selected worksheets before calling Workbook.Save with SaveFormat.Pdf.
// Common Searches: Aspose.Cells C# set BlackAndWhite page setup for specific sheets before PDF conversion | How to export confidential worksheets as black‑and‑white PDF with Aspose.Cells | Conditional page setup settings per worksheet Aspose.Cells .NET | Apply black and white printing mode to worksheets that contain a keyword using Aspose.Cells
// Tags: black-and-white page setup per worksheet Aspose.Cells | conditional worksheet printing settings .NET | export workbook to PDF with custom page setup | filter worksheets by name Aspose.Cells C# | apply printing mode before PDF conversion Aspose.Cells

using Aspose.Cells;
using System;

// Loads input.xlsx, enables black‑and‑white printing for any worksheet whose name includes “Confidential”, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Identify worksheets that contain confidential data.
            // Here we use a simple convention: worksheet name contains the word "Confidential".
            if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Apply black‑and‑white mode for printing/PDF rendering
                sheet.PageSetup.BlackAndWhite = true;
            }
        }

        // Export the workbook to PDF with the applied settings
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
