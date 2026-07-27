using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Rendering;

namespace SlicerPdfExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the slicer
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);

            // Convert the range into a table (ListObject)
            ListObject table = worksheet.ListObjects[worksheet.ListObjects.Add("A1", "B3", true)];

            // Add a slicer linked to the table's first column (Category)
            int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Validate the slicer's printable flag; set to true if it is false
            if (!slicer.IsPrintable)
            {
                slicer.IsPrintable = true;
                Console.WriteLine("Slicer.IsPrintable was false; set to true.");
            }
            else
            {
                Console.WriteLine("Slicer.IsPrintable is already true.");
            }

            // Prepare PDF save options (optional customizations can be added here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            string outputPath = "SlicerExported.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook exported to PDF at: {outputPath}");
        }
    }
}