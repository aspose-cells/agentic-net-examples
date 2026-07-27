using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;      // For Slicer class
using Aspose.Cells.Rendering;   // For PdfSaveOptions

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook that contains slicers
                string sourceFile = "input_with_slicers.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourceFile)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);

                // Ensure all slicers are printable so they appear in the PDF
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Slicer slicer in ws.Slicers)
                    {
                        slicer.IsPrintable = true;
                    }
                }

                // Configure PDF save options to retain document structure (includes slicer metadata)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook to PDF
                string outputFile = "output.pdf";
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF saved successfully: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}