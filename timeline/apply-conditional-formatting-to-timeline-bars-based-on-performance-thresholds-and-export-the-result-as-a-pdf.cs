// Title: Aspose.Cells C# – Apply Conditional Formatting to Timeline Bars and Export as PDF
// Description: This C# sample creates a workbook, fills it with date and performance data, builds a pivot table, attaches a timeline slicer to the Date field, applies three value‑based color rules to the Performance column, and saves the entire sheet—including the timeline and colored cells—as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# conditional formatting | timeline slicer PDF export | pivot table timeline Aspose | performance thresholds coloring | export to PDF Aspose.Cells | C# workbook PDF generation | Aspose.Cells timeline example | conditional formatting rules Aspose
// Common Searches: How to add a timeline slicer to a pivot table with Aspose.Cells .NET | Conditional formatting based on value ranges in Aspose.Cells C# | Export worksheet containing a timeline to PDF using Aspose.Cells | C# code for coloring performance cells and creating PDF report | Aspose.Cells example for KPI dashboard with timeline
// Developer Intent: Generate a PDF report that combines a pivot‑table timeline with value‑driven cell colors.
// Use Cases: Monthly KPI dashboard where users filter dates via a timeline and instantly see performance bands highlighted, then share as a printable PDF. | Automated sales performance summary that colors revenue cells, includes an interactive timeline, and outputs a PDF for email distribution. | Executive briefing requiring a timeline slicer for date selection and color‑coded metrics, saved as a PDF for archiving.
// AI Prompts: Modify the example to use custom threshold values and alternative colors for the conditional formatting. | Add data labels to the timeline bars while keeping the existing PDF export functionality. | Show how to programmatically resize and reposition the timeline shape based on worksheet dimensions. | Explain how to export only the timeline view as a separate PDF page. | Provide guidance on converting the generated PDF to an image while preserving conditional formatting.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This C# sample creates a workbook, fills it with date and performance data, builds a pivot table, attaches a timeline slicer to the Date field, applies three value‑based color rules to the Performance column, and saves the entire sheet—including the timeline and colored cells—as a PDF using Aspose.Cells for .NET.
class TimelineConditionalFormattingToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Date column (A) and Performance column (B)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Performance");

            // Sample dates and performance values
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 12; i++)
            {
                cells[i + 1, 0].PutValue(startDate.AddMonths(i));
                // Random performance between 30 and 100
                int perf = new Random(i).Next(30, 101);
                cells[i + 1, 1].PutValue(perf);
            }

            // Create a pivot table that uses the data range as its source
            string sourceRange = "A1:B13";
            string destCell = "D1";
            int pivotIdx = sheet.PivotTables.Add(sourceRange, destCell, "PerfPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            // Add Date to Row area and Performance to Data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Performance");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the pivot table on the Date field
            // Place the Timeline starting at row 15, column 0 (cell A15)
            int timelineIdx = sheet.Timelines.Add(pivot, 14, 0, "Date");
            // Optional: adjust size of the timeline shape
            sheet.Timelines[timelineIdx].Shape.Width = 500;
            sheet.Timelines[timelineIdx].Shape.Height = 80;

            // Apply conditional formatting to the Performance column (B2:B13)
            int cfIdx = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIdx];

            // Define the area to which the formatting will be applied
            CellArea perfArea = new CellArea
            {
                StartRow = 1,   // B2
                EndRow = 12,    // B13
                StartColumn = 1,
                EndColumn = 1
            };
            fcc.AddArea(perfArea);

            // Condition 1: Performance > 80 -> Green background
            int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "80", string.Empty);
            FormatCondition cond = fcc[condIdx];
            cond.Style.BackgroundColor = Color.LightGreen;

            // Condition 2: 50 <= Performance <= 80 -> Yellow background
            condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "50", "80");
            cond = fcc[condIdx];
            cond.Style.BackgroundColor = Color.LightYellow;

            // Condition 3: Performance < 50 -> Red background
            condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "50", string.Empty);
            cond = fcc[condIdx];
            cond.Style.BackgroundColor = Color.LightCoral;

            // Save the workbook as PDF (the timeline and conditional formatting will be rendered)
            workbook.Save("TimelineWithConditionalFormatting.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
