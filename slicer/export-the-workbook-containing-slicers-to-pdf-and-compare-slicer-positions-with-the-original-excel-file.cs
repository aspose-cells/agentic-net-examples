using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace SlicerPdfExportDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file that contains slicers
            string excelPath = "SlicersDemo.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(excelPath);

            // Assume slicers are on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Store slicer position information before saving
            var slicerInfo = new List<(int Index, double Left, double Top, double Width, double Height)>();

            foreach (Slicer slicer in worksheet.Slicers)
            {
                // Access the associated shape to get position and size
                var shape = slicer.Shape;

                slicerInfo.Add((
                    worksheet.Slicers.IndexOf(slicer),
                    shape.Left,
                    shape.Top,
                    shape.Width,
                    shape.Height));
            }

            // Set PDF save options (uses the provided save rule)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // retain document structure
            };

            // Export the workbook (including slicers) to PDF
            workbook.Save("SlicersExported.pdf", pdfOptions);

            // Compare and display slicer positions (they should match the original workbook)
            Console.WriteLine("Slicer positions in the original workbook:");
            foreach (var info in slicerInfo)
            {
                Console.WriteLine($"Slicer {info.Index}: Left={info.Left}, Top={info.Top}, Width={info.Width}, Height={info.Height}");
            }

            Console.WriteLine("Export to PDF completed.");
        }
    }
}