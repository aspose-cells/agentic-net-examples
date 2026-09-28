// Title: Validate slicer printable flag and enforce it before exporting a workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a slicer, checks its IsPrintable property, sets it to true when false, and then saves the workbook as a PDF with Aspose.Cells. | Show how to configure PdfSaveOptions while guaranteeing that all slicers in a workbook are marked printable before performing a PDF export in Aspose.Cells for .NET.
// Common Searches: aspnet cells how to verify slicer IsPrintable before PDF export | c# Aspose.Cells set slicer printable flag programmatically | export Excel workbook to PDF only when slicer printable is true | using PdfSaveOptions with slicer printable validation in Aspose.Cells | check slicer printable property Aspose.Cells before saving as PDF
// Tags: slicer IsPrintable property Aspose.Cells | set slicer printable flag C# | export workbook to PDF Aspose.Cells | PdfSaveOptions document structure Aspose.Cells | validate slicer before PDF export

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Rendering;

namespace SlicerPrintableCheckDemo
{
    // The example creates a workbook, adds a table and a slicer, verifies the slicer's IsPrintable flag, sets it to true if needed, and then saves the workbook as a PDF using PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data and create a table (required for slicer)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            // Add a ListObject (table) covering the data range
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Add a slicer linked to the table's first column
            int slicerIndex = sheet.Slicers.Add(table, 0, "D1");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Validate that the slicer is printable; if not, set it to true
            if (!slicer.IsPrintable)
            {
                // Optionally, you could throw an exception here
                // throw new InvalidOperationException("Slicer must be printable before exporting to PDF.");
                slicer.IsPrintable = true; // Ensure printable flag is true
            }

            // Prepare PDF save options (optional customizations)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: export document structure
                ExportDocumentStructure = true
            };

            // Export the workbook to PDF
            string outputPath = "SlicerPrintableChecked.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF at: {outputPath}");
        }
    }
}
