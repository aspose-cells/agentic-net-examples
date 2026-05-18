using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineConditionalFormattingPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Date and Performance values
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Performance");
            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 12; i++)
            {
                cells[i + 1, 0].PutValue(start.AddMonths(i));
                cells[i + 1, 1].PutValue(30 + i * 6);
            }

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIdx = sheet.PivotTables.Add("A1:B13", "D1", "PerfPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Performance");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table (based on the Date field)
            try
            {
                TimelineCollection timelines = sheet.Timelines;
                int timelineIdx = timelines.Add(pivot, "F1", "Date");
                Timeline timeline = timelines[timelineIdx];
                timeline.Shape.Width = 500;
                timeline.Shape.Height = 80;
                timeline.Caption = "Performance Timeline";
            }
            catch (Exception ex)
            {
                // Timeline creation can fail if the pivot configuration is not suitable.
                Console.WriteLine($"Timeline could not be added: {ex.Message}");
            }

            // Apply conditional formatting to the Performance column (B2:B13)
            int cfIdx = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIdx];

            // Define the range for conditional formatting
            CellArea perfArea = new CellArea
            {
                StartRow = 1,
                EndRow = 12,
                StartColumn = 1,
                EndColumn = 1
            };
            fcs.AddArea(perfArea);

            // Rule 1: Performance > 80 -> Green background
            int ruleIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "80", null);
            FormatCondition rule = fcs[ruleIdx];
            rule.Style.BackgroundColor = Color.LightGreen;

            // Rule 2: 50 <= Performance <= 80 -> Yellow background
            ruleIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "50", "80");
            rule = fcs[ruleIdx];
            rule.Style.BackgroundColor = Color.LightYellow;

            // Rule 3: Performance < 50 -> Red background
            ruleIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "50", null);
            rule = fcs[ruleIdx];
            rule.Style.BackgroundColor = Color.LightCoral;

            // Save the workbook as PDF (includes the timeline and conditional formatting)
            string outputPath = "PerformanceTimeline.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"An error occurred: {e.Message}");
        }
    }
}