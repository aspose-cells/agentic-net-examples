using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables; // For ListObject

namespace ExportWorkbookWithSlicers
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
                sheet.Name = "DataSheet";

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

                // Add a table (ListObject) covering the data range (A1:B5)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Header row is shown by default; totals row is not needed
                // (ShowHeader and ShowTotals properties are not available in this API version)

                // Add a slicer linked to the first column of the table
                // Position the slicer at cell D2 (row index 1, column index 3)
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], 1, 3);
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Shape.IsPrintable = true; // Ensure slicer is printed

                // Configure PDF save options to force all content onto a single page per sheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    AllColumnsInOnePagePerSheet = true
                };

                // Save the workbook as PDF; slicer will appear on the same page as the worksheet content
                string outputPath = "WorkbookWithSlicers.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}