// Title: C# Aspose.Cells: Convert Workbook to PDF with No Scaling (100% Zoom)
// Description: Demonstrates how to export an Aspose.Cells workbook to PDF while preserving the original worksheet size by enabling percent scaling and setting Zoom to 100%, eliminating any fit‑to‑page adjustments.
// Keywords: Aspose.Cells PDF conversion C# | no scaling PDF export | IsPercentScale true | Zoom 100 percent | preserve worksheet dimensions | disable fit-to-page | Excel to PDF Aspose | page setup PDF output | exact size PDF Aspose.Cells | C# workbook to PDF
// Common Searches: Aspose.Cells export to PDF without scaling | C# set page scaling none Aspose.Cells | How to keep original size when saving Excel as PDF | Disable fit-to-page in Aspose.Cells PDF conversion | 100% zoom PDF output Aspose.Cells
// Developer Intent: Generate a PDF from a workbook that matches the on‑screen layout by turning off automatic scaling.
// Use Cases: Create printable reports where the PDF must mirror the Excel view exactly. | Produce legal or regulatory documents that require precise table and chart dimensions. | Batch‑export multiple worksheets to PDF without automatic resizing.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as PDF with IsPercentScale = true and Zoom = 100. | Explain the effect of IsPercentScale and Zoom properties on PDF size in Aspose.Cells. | Show how to apply no‑scaling settings to every worksheet in a workbook before PDF conversion.

using System;
using Aspose.Cells;

// Demonstrates how to export an Aspose.Cells workbook to PDF while preserving the original worksheet size by enabling percent scaling and setting Zoom to 100%, eliminating any fit‑to‑page adjustments.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data (optional, just to have content)
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // ---- Set page scaling to NONE ----
        // Use 100% zoom and enable percent scaling so the worksheet is rendered
        // at its original size without any fit-to-page adjustments.
        sheet.PageSetup.IsPercentScale = true;   // Use percent scaling mode
        sheet.PageSetup.Zoom = 100;              // 100% zoom = no scaling

        // Save the workbook as PDF while preserving the original dimensions
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
