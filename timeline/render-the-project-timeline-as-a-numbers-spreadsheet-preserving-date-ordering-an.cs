using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ProjectTimelineToNumbers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ----- Populate data -----
        // Header row
        cells["A1"].Value = "Task";
        cells["B1"].Value = "StartDate";
        cells["C1"].Value = "EndDate";

        // Sample project tasks (intentionally unordered)
        cells["A2"].Value = "Design";
        cells["B2"].Value = new DateTime(2023, 5, 10);
        cells["C2"].Value = new DateTime(2023, 5, 20);

        cells["A3"].Value = "Development";
        cells["B3"].Value = new DateTime(2023, 5, 21);
        cells["C3"].Value = new DateTime(2023, 6, 30);

        cells["A4"].Value = "Testing";
        cells["B4"].Value = new DateTime(2023, 7, 1);
        cells["C4"].Value = new DateTime(2023, 7, 15);

        cells["A5"].Value = "Deployment";
        cells["B5"].Value = new DateTime(2023, 7, 16);
        cells["C5"].Value = new DateTime(2023, 7, 20);

        // ----- Apply date formatting -----
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        for (int r = 2; r <= 5; r++)
        {
            cells[r - 1, 1].SetStyle(dateStyle); // StartDate column (B)
            cells[r - 1, 2].SetStyle(dateStyle); // EndDate column (C)
        }

        // ----- Create a PivotTable to serve as the Timeline data source -----
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C5", "E1", "ProjectPivot");
        PivotTable pivot = pivots[pivotIndex];

        // Add fields: Task as rows, StartDate as columns, count of Task as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Task");
        pivot.AddFieldToArea(PivotFieldType.Column, "StartDate");
        pivot.AddFieldToArea(PivotFieldType.Data, "Task");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // ----- Add a Timeline linked to the PivotTable -----
        // Position the Timeline at row 15, column 0 (A16) and bind it to the "StartDate" field
        int timelineIndex = sheet.Timelines.Add(pivot, 15, 0, "StartDate");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Project Timeline";

        // ----- Save the workbook as an Apple Numbers file -----
        workbook.Save("ProjectTimeline.numbers", SaveFormat.Numbers);
    }
}