using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;      // For ListObject
using Aspose.Cells.Pivot;       // For PivotTable
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

class SlicerPdfReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Data";

            // Populate sample data that the slicer will be linked to
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Create a table (ListObject) from the data range
            int tableIdx = worksheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = worksheet.ListObjects[tableIdx];
            // Use DisplayName to set the table name (Name property not available in some versions)
            table.DisplayName = "DataTable";

            // Add a pivot table based on the created table
            int pivotIdx = worksheet.PivotTables.Add("DataTable", "D5", "PivotTable1");
            PivotTable pivot = worksheet.PivotTables[pivotIdx];

            // Add a slicer linked to the "Category" field of the pivot table
            int slicerIdx = worksheet.Slicers.Add(pivot, "Category", "CategorySlicer");
            Slicer slicer = worksheet.Slicers[slicerIdx];

            // Make the slicer printable
            slicer.Shape.IsPrintable = true;

            // Define a print area that includes the slicer (adjust as needed)
            string slicerRange = "C5:F10";
            worksheet.PageSetup.PrintArea = slicerRange;

            // Save the workbook as PDF; only the defined print area will be exported
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            string outputPath = "SlicerReport.pdf";

            // Ensure we can write the file (overwrite if it exists)
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}