using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelinePdfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date and numeric data
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Value";
            sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
            sheet.Cells["B3"].Value = 200;
            sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
            sheet.Cells["B4"].Value = 300;

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            // Configure PDF save options so the whole sheet (including the timeline) fits on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook as a PDF file
            workbook.Save("TimelineWorkbook.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}