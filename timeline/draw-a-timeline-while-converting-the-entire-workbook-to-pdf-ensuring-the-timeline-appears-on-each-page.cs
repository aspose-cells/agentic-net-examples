using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelinePdfDemo
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
                Cells cells = sheet.Cells;

                // Populate worksheet with sample data (date and value)
                cells["A1"].Value = "Date";
                cells["B1"].Value = "Value";

                cells["A2"].Value = new DateTime(2023, 1, 1);
                cells["B2"].Value = 1200;

                cells["A3"].Value = new DateTime(2023, 2, 1);
                cells["B3"].Value = 1500;

                cells["A4"].Value = new DateTime(2023, 3, 1);
                cells["B4"].Value = 1800;

                // Create a pivot table that will be the data source for the timeline
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table (placed at cell A10)
                // Use integer overload to avoid cell‑name parsing issues
                int timelineIdx = sheet.Timelines.Add(pivot, 9, 0, "Date"); // Row 9 (A10), Column 0 (A)
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Adjust the visual appearance of the timeline via its Shape object
                timeline.Shape.Top = 200;      // vertical offset in pixels
                timeline.Shape.Left = 50;      // horizontal offset in pixels
                timeline.Shape.Width = 600;    // width in pixels
                timeline.Shape.Height = 80;    // height in pixels
                timeline.Caption = "Sales Timeline";

                // Configure PDF save options – keep default pagination
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = false // allows multiple pages if the sheet is large
                };

                // Save the workbook as PDF; the timeline shape is part of the printable content
                workbook.Save("WorkbookWithTimeline.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}