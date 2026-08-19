// Title: Export Excel Slicers as Static Images in PDF using Aspose.Cells for .NET
// Description: Shows how to enable the IsPrintable flag for each slicer in a workbook and then save the file as PDF, resulting in slicers that are rendered as non‑interactive images while keeping their original look.
// Keywords: Aspose.Cells | C# | Excel slicer PDF | static slicer image | IsPrintable | export to PDF | non‑interactive slicer | save workbook as PDF | slicer rendering | PDF report generation
// Common Searches: Aspose.Cells render slicer as image PDF | make slicer non‑interactive in PDF C# | set slicer printable property Aspose | export Excel slicer to static PDF | PDF export slicer appearance Aspose.Cells
// Developer Intent: The developer needs to convert an Excel workbook to PDF while ensuring that slicers are displayed as fixed images rather than interactive controls.
// Use Cases: Create read‑only PDF dashboards that retain slicer visuals. | Archive Excel reports where slicer interactivity must be removed. | Automate batch conversion of workbooks to PDF with consistent slicer appearance.
// AI Prompts: Generate C# code with Aspose.Cells that exports a workbook to PDF and renders all slicers as static images. | Explain the effect of the Slicer.IsPrintable property on PDF output and how to apply it across a worksheet. | Recommend additional Aspose.Cells PDF settings to enhance the quality of slicer images in the final document.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Namespace for slicer objects

// Shows how to enable the IsPrintable flag for each slicer in a workbook and then save the file as PDF, resulting in slicers that are rendered as non‑interactive images while keeping their original look.
class RenderSlicerStaticPdf
{
    static void Main()
    {
        // Load the workbook that contains the slicer
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure each slicer is printable so it appears as a static image in the PDF
        foreach (Slicer slicer in worksheet.Slicers)
        {
            slicer.IsPrintable = true;
        }

        // Save the workbook as PDF; slicers will be rendered as non‑interactive images
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
