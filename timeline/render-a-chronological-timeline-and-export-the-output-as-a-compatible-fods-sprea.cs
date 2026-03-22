using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineToFodsDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate sample data with dates (chronological order) and values
            // Header row
            cells[0, 0].Value = "Event";
            cells[0, 1].Value = "Date";
            cells[0, 2].Value = "Amount";

            // Sample rows
            cells[1, 0].Value = "Project Kickoff";
            cells[1, 1].Value = new DateTime(2021, 1, 15);
            cells[1, 2].Value = 1000;

            cells[2, 0].Value = "Phase 1 Complete";
            cells[2, 1].Value = new DateTime(2021, 4, 10);
            cells[2, 2].Value = 2500;

            cells[3, 0].Value = "Phase 2 Complete";
            cells[3, 1].Value = new DateTime(2021, 7, 5);
            cells[3, 2].Value = 4000;

            cells[4, 0].Value = "Final Delivery";
            cells[4, 1].Value = new DateTime(2021, 12, 20);
            cells[4, 2].Value = 6000;

            // Apply a date style to the Date column for proper display
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "yyyy-mm-dd";
            for (int row = 1; row <= 4; row++)
            {
                cells[row, 1].SetStyle(dateStyle);
            }

            // 3. Create a PivotTable based on the data range
            // Data range: A1:C5
            // Place the PivotTable at cell E3
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "TimelinePivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields: Event as Row, Date as Column (time axis), Amount as Data
            pivot.AddFieldToArea(PivotFieldType.Row, "Event");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh and calculate the PivotTable
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline control linked to the PivotTable's Date field
            // Place the Timeline starting at cell G1 (row 0, column 6)
            int timelineIndex = sheet.Timelines.Add(pivot, "G1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set a caption for better readability
            timeline.Caption = "Project Timeline";

            // 5. Save the workbook as an OpenDocument Flat XML Spreadsheet (FODS)
            workbook.Save("ProjectTimeline.fods", SaveFormat.Fods);
        }
    }
}