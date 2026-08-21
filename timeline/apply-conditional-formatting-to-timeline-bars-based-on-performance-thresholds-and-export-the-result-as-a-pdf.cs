// Title: Conditional Formatting on Aspose.Cells Timeline Bars with Performance Thresholds and PDF Export (C#)
// Description: Demonstrates how to create a workbook with date‑performance data, build a pivot table, attach a timeline, apply three value‑based conditional formatting rules (green > 80, yellow > 50, red ≤ 50) to the performance column, and save the result as a PDF using Aspose.Cells for .NET. Ideal for global reporting scenarios.
// Keywords: Aspose.Cells | C# | timeline | conditional formatting | pivot table | PDF export | performance thresholds | color coding | Aspose.Cells for .NET | timeline bars
// Common Searches: Aspose.Cells timeline conditional formatting C# | Export timeline to PDF with Aspose.Cells | Apply color rules to cells based on value Aspose.Cells | Create pivot table and timeline programmatically | How to color timeline bars by performance in .NET
// Developer Intent: Generate a workbook, link a timeline to a pivot table, color‑code performance values, and output the formatted timeline as a PDF.
// Use Cases: Project status reports that highlight low, medium, and high performance periods. | Automated monthly dashboards with colored timeline bars for quick visual analysis. | PDF‑based performance summaries for stakeholders that require clear visual cues.
// AI Prompts: Write C# code with Aspose.Cells to add a timeline to a pivot table, apply three conditional formatting rules (green > 80, yellow > 50, red ≤ 50) on column B, and export the workbook as a PDF. | Show how to set background colors for cells B2:B6 based on numeric thresholds using Aspose.Cells conditional formatting. | Explain how to modify the timeline caption and style after applying conditional formatting in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with date‑performance data, build a pivot table, attach a timeline, apply three value‑based conditional formatting rules (green > 80, yellow > 50, red ≤ 50) to the performance column, and save the result as a PDF using Aspose.Cells for .NET. Ideal for global reporting scenarios.
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

                // Populate sample data: Date and Performance columns
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Performance");

                // Add several rows of data
                cells["A2"].PutValue(new DateTime(2023, 1, 1));
                cells["B2"].PutValue(45);   // Low performance
                cells["A3"].PutValue(new DateTime(2023, 2, 1));
                cells["B3"].PutValue(65);   // Medium performance
                cells["A4"].PutValue(new DateTime(2023, 3, 1));
                cells["B4"].PutValue(85);   // High performance
                cells["A5"].PutValue(new DateTime(2023, 4, 1));
                cells["B5"].PutValue(55);   // Medium performance
                cells["A6"].PutValue(new DateTime(2023, 5, 1));
                cells["B6"].PutValue(30);   // Low performance

                // Create a pivot table that will serve as the data source for the timeline
                int pivotIdx = sheet.PivotTables.Add("A1:B6", "D1", "PerformancePivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Performance");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table (based on the Date field)
                TimelineCollection timelines = sheet.Timelines;
                int timelineIdx = timelines.Add(pivot, "F1", "Date");
                Timeline timeline = timelines[timelineIdx];
                timeline.Caption = "Performance Timeline";

                // Apply conditional formatting to the Performance column based on thresholds
                int cfIdx = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIdx];

                // Define the range to which the formatting will be applied (B2:B6)
                CellArea area = new CellArea
                {
                    StartRow = 1,
                    EndRow = 5,
                    StartColumn = 1,
                    EndColumn = 1
                };
                fcs.AddArea(area);

                // High performance (> 80) – green background
                int highIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "80", null);
                FormatCondition highCond = fcs[highIdx];
                highCond.Style.BackgroundColor = Color.LightGreen;

                // Medium performance (> 50) – yellow background
                int mediumIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
                FormatCondition mediumCond = fcs[mediumIdx];
                mediumCond.Style.BackgroundColor = Color.LightYellow;

                // Low performance (<= 50) – red background
                // Use LessThan with a threshold of 51 to emulate <= 50
                int lowIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "51", null);
                FormatCondition lowCond = fcs[lowIdx];
                lowCond.Style.BackgroundColor = Color.LightCoral;

                // Save the workbook as a PDF file
                workbook.Save("TimelineConditionalFormatting.pdf", SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
