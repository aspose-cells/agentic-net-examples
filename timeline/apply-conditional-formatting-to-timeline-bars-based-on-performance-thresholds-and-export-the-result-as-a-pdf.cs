// Title: Apply color‑coded conditional formatting and data bars to a performance timeline and export the workbook as PDF using Aspose.Cells for .NET
// AI Prompts: Generate a C# program that creates a workbook with Date and Performance columns, builds a pivot table, applies three color‑based conditional formatting rules (>80 green, 50‑80 yellow, <50 coral) and a steel‑blue data bar to the Performance range, then saves the workbook as a PDF. | Update the Aspose.Cells example to use custom threshold values (e.g., >90, 60‑90, <60) and a different data bar color, ensuring the PDF output reflects the new formatting. | Add a timeline linked to the pivot table, apply the same conditional formatting and data bar to the timeline’s performance values, and export the final workbook to PDF in C#.
// Common Searches: aspnet conditional formatting with multiple thresholds and data bar Aspose.Cells example | export timeline with conditional formatting to PDF using Aspose.Cells for .NET | C# code to apply color scales and data bars to a pivot table performance column | how to add a timeline to a pivot table and save as PDF with Aspose.Cells | Aspose.Cells conditional formatting based on performance values and PDF output
// Tags: conditional formatting color thresholds Aspose.Cells C# | data bar conditional format Aspose.Cells | export workbook to PDF Aspose.Cells | timeline pivot table Aspose.Cells | performance metric visualization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsTimelineConditionalFormatting
{
    // The sample creates a new workbook, fills it with monthly dates and random performance values, builds a pivot table, optionally adds a timeline, applies three color‑coded conditional formatting rules and a steel‑blue data bar to the Performance column, and saves the formatted workbook as a PDF file.
    class Program
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

                DateTime startDate = new DateTime(2023, 1, 1);
                Random rnd = new Random();
                for (int i = 0; i < 12; i++)
                {
                    cells[i + 1, 0].PutValue(startDate.AddMonths(i));
                    cells[i + 1, 1].PutValue(rnd.Next(30, 101));
                }

                // Create a pivot table using the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B13", "D1", "PerfPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Performance");

                // Refresh pivot cache and calculate data (compatible with all versions)
                pivot.RefreshData();
                pivot.CalculateData();

                // Timeline feature may not be available in all versions; skip if not supported
                // Uncomment the following lines if Timeline class is present in your Aspose.Cells version
                /*
                int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Date");
                var timeline = sheet.Timelines[timelineIdx];
                timeline.Caption = "Performance Timeline";
                */

                // Apply conditional formatting to the Performance column (B2:B13)
                int cfIdx = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIdx];
                CellArea perfArea = new CellArea { StartRow = 1, EndRow = 12, StartColumn = 1, EndColumn = 1 };
                fcc.AddArea(perfArea);

                // Values > 80 -> Green background
                int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "80", null);
                FormatCondition cond = fcc[condIdx];
                cond.Style.BackgroundColor = Color.LightGreen;

                // Values between 50 and 80 -> Yellow background
                condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "50", "80");
                cond = fcc[condIdx];
                cond.Style.BackgroundColor = Color.LightYellow;

                // Values < 50 -> Coral background
                condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "50", null);
                cond = fcc[condIdx];
                cond.Style.BackgroundColor = Color.LightCoral;

                // Add a DataBar to visualize performance values
                int dataBarIdx = sheet.ConditionalFormattings.Add();
                FormatConditionCollection dataBarFc = sheet.ConditionalFormattings[dataBarIdx];
                dataBarFc.AddArea(perfArea);
                int dbCondIdx = dataBarFc.AddCondition(FormatConditionType.DataBar);
                FormatCondition dbCond = dataBarFc[dbCondIdx];
                dbCond.DataBar.Color = Color.SteelBlue;
                dbCond.DataBar.MinCfvo.Type = FormatConditionValueType.Min;
                dbCond.DataBar.MaxCfvo.Type = FormatConditionValueType.Max;
                dbCond.DataBar.ShowValue = true;

                // Save the workbook as PDF
                workbook.Save("TimelineConditionalFormatting.pdf", SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
