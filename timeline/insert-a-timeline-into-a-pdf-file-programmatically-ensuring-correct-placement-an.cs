using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data (Date + Value)
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Value";

            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 120;

            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 150;

            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 180;

            // Ensure the date column is formatted as a date
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Short date format
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            cells.CreateRange("A2:A4").ApplyStyle(dateStyle, flag);

            // Create a PivotTable that will be the data source for the Timeline
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Add the Date field to the Page area (required for Timeline) and Value as data
            pivot.AddFieldToArea(PivotFieldType.Page, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (destination cell F5 -> row 4, column 5)
            int timelineIdx = sheet.Timelines.Add(pivot, 4, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];

            // Fine‑tune visual placement using the underlying shape
            timeline.Shape.Left = 100;   // pixels from the left edge of the worksheet
            timeline.Shape.Top = 50;     // pixels from the top edge of the worksheet
            timeline.Shape.Width = 400;  // width in pixels
            timeline.Shape.Height = 80;  // height in pixels
            timeline.Caption = "Sales Timeline";

            // Save the workbook as a PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedStandardWindowsFonts = true
            };

            workbook.Save("TimelineOutput.pdf", pdfOptions);
        }
    }
}