using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

class ExportSlicersToPdf
{
    static void Main()
    {
        // Load the workbook that contains slicers
        string excelPath = "SlicersDemo.xlsx";
        Workbook workbook = new Workbook(excelPath);

        // Collect slicer position information from the original workbook
        var slicerInfo = new List<string>();
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Slicer slicer in ws.Slicers)
            {
                // Access the shape associated with the slicer
                var shape = slicer.Shape;

                // Record basic position and size properties
                slicerInfo.Add(
                    $"Worksheet: {ws.Name}, Slicer: {slicer.Name}, " +
                    $"Top={shape.Top}, Left={shape.Left}, " +
                    $"Width={shape.Width}, Height={shape.Height}");
            }
        }

        // Save the workbook to PDF, exporting the document structure
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // retain document structure in PDF
        workbook.Save("SlicersOutput.pdf", pdfOptions);

        // Output the collected slicer positions for comparison
        Console.WriteLine("Slicer positions in the original Excel file:");
        foreach (string info in slicerInfo)
        {
            Console.WriteLine(info);
        }
    }
}