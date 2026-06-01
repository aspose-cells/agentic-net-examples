using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data to create a table (required for slicer)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");

        // Add a ListObject (table) covering the data range
        ListObject table = worksheet.ListObjects[worksheet.ListObjects.Add("A1", "A3", true)];

        // Add a slicer linked to the table at cell D1
        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Validate that the slicer is printable; set to true if not
        if (!slicer.IsPrintable)
        {
            slicer.IsPrintable = true;
        }

        // Configure PDF save options (default options are sufficient for this example)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}