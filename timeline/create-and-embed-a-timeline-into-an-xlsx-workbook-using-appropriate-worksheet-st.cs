using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample date and sales data
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Sales";

        cells["A2"].Value = new DateTime(2023, 1, 1);
        cells["B2"].Value = 1000;
        cells["A3"].Value = new DateTime(2023, 1, 15);
        cells["B3"].Value = 1500;
        cells["A4"].Value = new DateTime(2023, 2, 1);
        cells["B4"].Value = 2000;
        cells["A5"].Value = new DateTime(2023, 2, 15);
        cells["B5"].Value = 2500;

        // Apply a date format to the date column
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        cells["A2"].SetStyle(dateStyle);
        cells["A3"].SetStyle(dateStyle);
        cells["A4"].SetStyle(dateStyle);
        cells["A5"].SetStyle(dateStyle);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table.
        // The timeline will be placed with its upper‑left corner at row 10, column 0 (cell A11)
        int timelineIndex = sheet.Timelines.Add(pivot, 10, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];

        // Optional: customize the timeline appearance via its Shape object
        timeline.Caption = "Sales Timeline";
        timeline.Shape.Width = 400;   // width in pixels
        timeline.Shape.Height = 120;  // height in pixels
        timeline.Shape.Left = 50;     // left offset in pixels
        timeline.Shape.Top = 200;     // top offset in pixels

        // Save the workbook to an XLSX file
        workbook.Save("TimelineExample.xlsx");
    }
}