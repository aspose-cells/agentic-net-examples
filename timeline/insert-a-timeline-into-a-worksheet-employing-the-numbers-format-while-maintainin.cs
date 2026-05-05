using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class InsertTimelineNumbersFormat
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample data (Date and numeric Amount)
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Amount");

        cells["A2"].PutValue(new DateTime(2023, 1, 1));
        cells["B2"].PutValue(150);
        cells["A3"].PutValue(new DateTime(2023, 2, 1));
        cells["B3"].PutValue(200);
        cells["A4"].PutValue(new DateTime(2023, 3, 1));
        cells["B4"].PutValue(250);

        // Apply a numeric format to the Amount column (NUMBERS format)
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Custom = "0"; // integer format
        cells["B2"].SetStyle(numberStyle);
        cells["B3"].SetStyle(numberStyle);
        cells["B4"].SetStyle(numberStyle);

        // Apply a date format to the Date column (optional, keeps data readable)
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        cells["A2"].SetStyle(dateStyle);
        cells["A3"].SetStyle(dateStyle);
        cells["A4"].SetStyle(dateStyle);

        // Create a pivot table based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIdx = pivots.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = pivots[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table, positioned at cell E5
        int timelineIdx = sheet.Timelines.Add(pivot, "E5", "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];
        timeline.Caption = "Sales Timeline";

        // Save the workbook
        workbook.Save("TimelineNumbersFormat.xlsx");
    }
}