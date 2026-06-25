using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

class TimelineToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate data ----------
            // Header
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            // Add dates with a weekly interval and some sample values
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 12; i++)
            {
                // Date column (weekly)
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 7));

                // Value column (any numeric data)
                sheet.Cells[i + 1, 1].PutValue(i * 10);
            }

            // ---------- Create a PivotTable ----------
            // Use the data range A1:B13 (header + 12 rows)
            int pivotIndex = sheet.PivotTables.Add("A1:B13", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Row field = Date, Data field = Value
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a Timeline linked to the PivotTable ----------
            // Place the timeline starting at row 0, column 5 (cell F1)
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Set a caption for the timeline
            timeline.Caption = "Weekly Timeline";

            // Apply bold formatting to the caption via the underlying shape's font
            timeline.Shape.Font.IsBold = true;

            // ---------- Save the workbook as PDF ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("Timeline.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}