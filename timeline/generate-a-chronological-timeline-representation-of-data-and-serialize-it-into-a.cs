using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

class TimelineSpreadsheetMLDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Event, Date, Value
        cells["A1"].PutValue("Event");
        cells["B1"].PutValue("Date");
        cells["C1"].PutValue("Value");

        string[] events = { "Launch", "Update", "Release", "Anniversary", "Retirement" };
        DateTime[] dates = {
            new DateTime(2020, 1, 15),
            new DateTime(2021, 5, 10),
            new DateTime(2022, 9, 30),
            new DateTime(2023, 12, 1),
            new DateTime(2024, 6, 20)
        };
        int[] values = { 100, 150, 200, 250, 300 };

        for (int i = 0; i < events.Length; i++)
        {
            cells[i + 1, 0].PutValue(events[i]);          // Event column
            cells[i + 1, 1].PutValue(dates[i]);           // Date column
            cells[i + 1, 2].PutValue(values[i]);          // Value column
        }

        // Apply a date style to the Date column
        Style dateStyle = new CellsFactory().CreateStyle();
        dateStyle.Custom = "yyyy-mm-dd";
        for (int i = 1; i <= events.Length; i++)
        {
            cells[i, 1].SetStyle(dateStyle);
        }

        // Create a PivotTable based on the data range
        PivotTableCollection pivotTables = sheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:C6", "E3", "TimelinePivot");
        PivotTable pivot = pivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Event");
        pivot.AddFieldToArea(PivotFieldType.Column, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable using the Date field
        TimelineCollection timelines = sheet.Timelines;
        int timelineIndex = timelines.Add(pivot, "A10", "Date");
        Timeline timeline = timelines[timelineIndex];
        timeline.Caption = "Project Timeline";

        // Save the workbook as SpreadsheetML (Excel 2003 XML)
        SpreadsheetML2003SaveOptions saveOptions = new SpreadsheetML2003SaveOptions();
        workbook.Save("TimelineSpreadsheetML.xml", saveOptions);
    }
}