// Title: Convert Excel slicers to static images when saving as PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through all worksheets, sets each slicer's IsPrintable property to true, and saves the workbook as a PDF using Aspose.Cells. | Show how to ensure slicers are rendered as non‑interactive images in a PDF export with Aspose.Cells for .NET. | Provide a snippet that disables slicer interactivity by marking slicers printable before calling Workbook.Save in PDF format.
// Common Searches: Aspose.Cells how to export slicers as images in PDF C# | C# make Excel slicer non‑clickable in PDF output using Aspose | set slicer printable flag before saving workbook to PDF with Aspose.Cells | render Excel slicer as static picture in PDF conversion .NET
// Tags: Aspose.Cells set slicer printable | export slicer to PDF static image | C# disable slicer interactivity Aspose | PDF conversion slicer rendering Aspose.Cells | Excel slicer printable property .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Namespace containing the Slicer class

// Loads an Excel workbook, marks every slicer as printable, and saves the workbook as a PDF so slicers appear as static, non‑interactive images.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure every slicer is marked as printable so it will be rendered in the PDF
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Slicer slicer in sheet.Slicers)
            {
                // The IsPrintable property (though obsolete) controls slicer visibility in print/PDF output
                slicer.IsPrintable = true;
            }
        }

        // Save the workbook as a PDF file; slicers will now appear as static images
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
