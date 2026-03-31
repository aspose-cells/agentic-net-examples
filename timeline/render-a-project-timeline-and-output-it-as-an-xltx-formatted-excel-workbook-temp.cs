using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ProjectTimelineTemplate
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample project data: Task, StartDate, EndDate, Owner
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("StartDate");
        sheet.Cells["C1"].PutValue("EndDate");
        sheet.Cells["D1"].PutValue("Owner");

        sheet.Cells["A2"].PutValue("Planning");
        sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 15));
        sheet.Cells["D2"].PutValue("Alice");

        sheet.Cells["A3"].PutValue("Design");
        sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 16));
        sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 5));
        sheet.Cells["D3"].PutValue("Bob");

        sheet.Cells["A4"].PutValue("Implementation");
        sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 6));
        sheet.Cells["C4"].PutValue(new DateTime(2023, 4, 30));
        sheet.Cells["D4"].PutValue("Charlie");

        sheet.Cells["A5"].PutValue("Testing");
        sheet.Cells["B5"].PutValue(new DateTime(2023, 5, 1));
        sheet.Cells["C5"].PutValue(new DateTime(2023, 5, 31));
        sheet.Cells["D5"].PutValue("Diana");

        // Apply a date style to the date columns
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        for (int row = 1; row <= 5; row++)
        {
            sheet.Cells[row, 1].SetStyle(dateStyle); // StartDate column (B)
            sheet.Cells[row, 2].SetStyle(dateStyle); // EndDate column (C)
        }

        // Create a PivotTable that uses StartDate as a column field (required for Timeline)
        PivotTableCollection pivots = sheet.PivotTables;
        // The data range includes all columns; the pivot will be placed at G3
        int pivotIndex = pivots.Add("A1:D5", "G3", "ProjectPivot");
        PivotTable pivot = pivots[pivotIndex];

        // Add fields: Task as Row, StartDate as Column, Owner as Data (count)
        pivot.AddFieldToArea(PivotFieldType.Row, "Task");
        pivot.AddFieldToArea(PivotFieldType.Column, "StartDate");
        pivot.AddFieldToArea(PivotFieldType.Data, "Owner");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

        // Refresh and calculate the PivotTable
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable using the StartDate field
        // Position the Timeline at row 15, column 0 (A16)
        int timelineIndex = sheet.Timelines.Add(pivot, 15, 0, "StartDate");
        Timeline timeline = sheet.Timelines[timelineIndex];

        // Optional: customize Timeline appearance
        timeline.Caption = "Project Timeline";
        timeline.ShowHeader = true;
        timeline.ShowHorizontalScrollbar = true;
        timeline.ShowSelectionLabel = true;
        timeline.ShowTimeLevel = true;

        // Save the workbook as an Excel template (XLTX) using SaveOptions (lifecycle: save)
        // Using XlsSaveOptions with IsTemplate = true (obsolete but provided in rules)
        XlsSaveOptions saveOptions = new XlsSaveOptions();
        saveOptions.IsTemplate = true; // marks the file as a template

        using (FileStream fs = File.Create("ProjectTimelineTemplate.xlt"))
        {
            workbook.Save(fs, saveOptions);
        }

        // The template file "ProjectTimelineTemplate.xlt" now contains the timeline control.
    }
}