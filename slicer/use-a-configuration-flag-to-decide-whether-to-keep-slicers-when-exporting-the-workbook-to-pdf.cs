using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerPdfExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Configuration flag: true to keep slicers in the PDF, false to hide them
                bool keepSlicersInPdf = GetKeepSlicersFlag();

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the slicer
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["A5"].PutValue("B");

                // Convert the range to a table (ListObject) so that a slicer can be attached
                int tableIndex = sheet.ListObjects.Add("A1", "A5", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set a display name for the table (Name property may not be available in some versions)
                table.DisplayName = "CategoryTable";

                // Add a slicer linked to the table (column index 0 corresponds to the first column)
                int slicerIndex = sheet.Slicers.Add(table, 0, "D1");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Control slicer visibility in the PDF using Shape.IsPrintable
                slicer.Shape.IsPrintable = keepSlicersInPdf;

                // Configure PDF save options (default settings are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook to PDF
                string outputPath = "ExportedWorkbook.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook exported successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Placeholder for obtaining the configuration flag (could be from app settings, env variable, etc.)
        static bool GetKeepSlicersFlag()
        {
            // For demonstration, we simply return true.
            // Replace this logic with actual configuration retrieval as needed.
            return true;
        }
    }
}