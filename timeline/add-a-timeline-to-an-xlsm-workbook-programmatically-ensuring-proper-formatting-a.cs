using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class AddTimelineToMacroWorkbook
{
    static void Main()
    {
        // Create a new workbook (macro-enabled will be saved as .xlsm)
        Workbook workbook = new Workbook();

        // Enable macros in workbook settings (optional, ensures macro support)
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a date column (required for timeline)
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Date");
        cells["C1"].PutValue("Sales");

        cells["A2"].PutValue("P1");
        cells["B2"].PutValue(new DateTime(2023, 1, 1));
        cells["C2"].PutValue(120);

        cells["A3"].PutValue("P2");
        cells["B3"].PutValue(new DateTime(2023, 1, 5));
        cells["C3"].PutValue(150);

        cells["A4"].PutValue("P1");
        cells["B4"].PutValue(new DateTime(2023, 1, 10));
        cells["C4"].PutValue(200);

        // Apply date format to the Date column
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        for (int row = 1; row <= 4; row++)
        {
            cells[row, 1].SetStyle(dateStyle);
        }

        // Create a pivot table that will serve as the timeline data source
        int pivotIdx = sheet.PivotTables.Add("A1:C4", "E1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Column, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table (placed at cell A10)
        int timelineIdx = sheet.Timelines.Add(pivot, "A10", "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];

        // Configure timeline appearance
        timeline.Caption = "Sales Timeline";
        timeline.ShowHeader = true;
        timeline.ShowHorizontalScrollbar = true;
        timeline.ShowSelectionLabel = true;
        timeline.ShowTimeLevel = true;

        // Optionally adjust size via the associated shape
        timeline.Shape.Width = 500;   // width in pixels
        timeline.Shape.Height = 120;  // height in pixels

        // Save the workbook as a macro-enabled file
        workbook.Save("TimelineWithMacro.xlsm", SaveFormat.Xlsm);
    }
}