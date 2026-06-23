using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate data with weekly dates -----
            // Header row
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");

            // Apply bold font to header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            cells["A1"].SetStyle(headerStyle);
            cells["B1"].SetStyle(headerStyle);

            // Date style (so Aspose.Cells recognises the values as dates)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // built‑in date format

            // Fill 8 weeks of data (weekly interval)
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 8; i++)
            {
                // Date column (weekly interval)
                Cell dateCell = cells[i + 1, 0];
                dateCell.PutValue(startDate.AddDays(i * 7));
                dateCell.SetStyle(dateStyle);

                // Sample sales value
                cells[i + 1, 1].PutValue(1000 + i * 200);
            }

            // ----- Create a PivotTable based on the data -----
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:B9", "D1", "SalesPivot");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // ----- Add a Timeline linked to the PivotTable -----
            // Place the Timeline starting at row 0, column 5 (cell F1)
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Weekly Sales Timeline";

            // Adjust the visual size of the Timeline
            timeline.Shape.Width = 400;   // width in points
            timeline.Shape.Height = 100;  // height in points

            // ----- Prepare PDF save options with a bold watermark -----
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Create a bold rendering font for the watermark
            RenderingFont watermarkFont = new RenderingFont("Arial", 24)
            {
                Bold = true,
                Color = Color.DarkBlue
            };
            pdfOptions.Watermark = new RenderingWatermark("Generated PDF", watermarkFont);

            // ----- Save the workbook (including the Timeline) as a PDF document -----
            string outputPath = "TimelineWeekly.pdf";
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}