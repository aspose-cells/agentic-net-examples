using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ExportTimelinePdfToMemoryStream
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date and value data
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Value";

            sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
            sheet.Cells["B2"].Value = 100;

            sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
            sheet.Cells["B3"].Value = 150;

            sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
            sheet.Cells["B4"].Value = 200;

            // Add a pivot table that will serve as the data source for the timeline
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table's date field
            // Place the timeline starting at cell G1 (row 0, column 6) to avoid overlap
            sheet.Timelines.Add(pivot, 0, 6, "Date");

            // Export the workbook (which now contains the timeline) to a PDF in a memory stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0;
                Console.WriteLine($"PDF stream length: {pdfStream.Length} bytes");
                // The pdfStream can now be used for further processing without any disk I/O
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}