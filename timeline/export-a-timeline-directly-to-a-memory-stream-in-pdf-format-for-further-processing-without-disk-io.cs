using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ExportTimelineToPdfStream
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date and sales data
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Sales";

            sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
            sheet.Cells["B2"].Value = 100;

            sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
            sheet.Cells["B3"].Value = 150;

            sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
            sheet.Cells["B4"].Value = 200;

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();          // Refresh source data
            pivot.CalculateData();        // Calculate pivot values

            // Add a timeline linked to the pivot table (date field must be in the row area)
            int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];
            timeline.Caption = "Sales Timeline";

            // Save the workbook (including the timeline) to PDF in a memory stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0; // Reset for further processing

                Console.WriteLine($"PDF stream length: {pdfStream.Length} bytes");
                // pdfStream can now be transmitted, stored, etc.
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}