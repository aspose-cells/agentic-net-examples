using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;      // For ListObject
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;   // For PdfSaveOptions

namespace SlicerPdfExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("C");
                sheet.Cells["B5"].PutValue(40);

                // Add a table that covers the data range (A1:B5)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "DataTable"; // Friendly name

                // Add a slicer linked to the first column of the table (Category)
                // Position the slicer at D2 (row 1, column 3)
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], 1, 3);
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Shape.IsPrintable = true; // Ensure slicer appears in PDF

                // Configure PDF save options to fit all content on a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    AllColumnsInOnePagePerSheet = true
                };

                // Define output file path and ensure directory exists
                string outputPath = "SlicersExport.pdf";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF; slicer will appear on the same page as the data
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}