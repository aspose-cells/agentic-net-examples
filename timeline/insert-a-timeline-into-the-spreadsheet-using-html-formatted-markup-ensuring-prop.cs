using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

class InsertTimelineWithHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a date field
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Value";

        cells["A2"].Value = new DateTime(2023, 1, 1);
        Style styleA2 = cells["A2"].GetStyle();
        styleA2.Number = 14; // Date format
        cells["A2"].SetStyle(styleA2);
        cells["B2"].Value = 100;

        cells["A3"].Value = new DateTime(2023, 2, 1);
        Style styleA3 = cells["A3"].GetStyle();
        styleA3.Number = 14;
        cells["A3"].SetStyle(styleA3);
        cells["B3"].Value = 200;

        cells["A4"].Value = new DateTime(2023, 3, 1);
        Style styleA4 = cells["A4"].GetStyle();
        styleA4.Number = 14;
        cells["A4"].SetStyle(styleA4);
        cells["B4"].Value = 300;

        // Create a pivot table that will serve as the timeline data source
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table (using the date field)
        int timelineIdx = sheet.Timelines.Add(pivot, 5, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];

        // Insert HTML-formatted text into the timeline shape
        timeline.Shape.HtmlText = "<div style='font-weight:bold;color:#0066CC;'>Sales Timeline</div>";

        // Adjust size and position using Shape properties (optional)
        timeline.Shape.Width = 400;
        timeline.Shape.Height = 120;
        timeline.Shape.Top = 20;
        timeline.Shape.Left = 20;

        // Save the workbook as an Excel file
        workbook.Save("TimelineWithHtml.xlsx");

        // Save as HTML preserving the HTML markup inside the timeline shape
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ParseHtmlTagInCell = true;
        htmlOptions.SaveAsSingleFile = true;
        workbook.Save("TimelineWithHtml.html", htmlOptions);
    }
}