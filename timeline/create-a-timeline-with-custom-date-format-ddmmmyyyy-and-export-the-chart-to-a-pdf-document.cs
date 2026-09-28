// Title: Generate a timeline with dd-MMM-yyyy date format, link it to a pivot table, create a line chart, and export the chart as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a timeline with a custom dd-MMM-yyyy date format linked to a pivot table, and saves the resulting chart as a PDF using Aspose.Cells. | Show how to apply a custom date style to worksheet cells and a chart’s X‑axis, then export the chart to a PDF file in Aspose.Cells for .NET. | Provide a complete example that builds sample sales data, constructs a pivot table, inserts a timeline, creates a line chart using the timeline dates, and outputs the chart to a PDF.
// Common Searches: Aspose.Cells C# timeline custom date format PDF export | how to set dd-MMM-yyyy format on timeline and chart axis in Aspose.Cells | export line chart linked to a pivot table as PDF using Aspose.Cells .NET | create timeline control for pivot table and save chart to PDF in C# | apply custom date style to chart X axis Aspose.Cells example
// Tags: Aspose.Cells timeline custom date format | export chart to PDF Aspose.Cells C# | pivot table timeline integration Aspose.Cells | set chart X axis date format Aspose.Cells | line chart PDF generation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelinePdfDemo
{
    // Demonstrates creating a workbook with sample sales data, applying a dd-MMM-yyyy custom date style, building a pivot table, inserting a linked timeline, generating a line chart that uses the timeline dates, and exporting the chart to a PDF file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Populate sample data ----------
            // Header row
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Date";
            cells["C1"].Value = "Sales";

            // Data rows
            cells["A2"].Value = "P1";
            cells["A3"].Value = "P2";
            cells["A4"].Value = "P3";
            cells["A5"].Value = "P4";

            DateTime[] dates = {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 2, 12),
                new DateTime(2023, 3, 20),
                new DateTime(2023, 4, 28)
            };
            for (int i = 0; i < dates.Length; i++)
            {
                cells[i + 2, 1].Value = dates[i]; // Column B (index 1)
            }

            cells["C2"].Value = 120;
            cells["C3"].Value = 150;
            cells["C4"].Value = 180;
            cells["C5"].Value = 210;

            // ---------- Apply custom date format (dd-MMM-yyyy) ----------
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            for (int i = 0; i < dates.Length; i++)
            {
                cells[i + 2, 1].SetStyle(dateStyle); // Apply to each date cell
            }

            // ---------- Create a PivotTable ----------
            // Data range A1:C5, destination cell E1
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            // Add fields to the pivot
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            // Refresh to calculate data
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a Timeline linked to the PivotTable ----------
            // Place the Timeline starting at cell G1, using the "Date" base field
            int timelineIndex = sheet.Timelines.Add(pivot, "G1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            // Optional: set a caption for the Timeline
            timeline.Caption = "Sales Timeline";

            // ---------- Create a Chart that uses the same data ----------
            int chartIndex = sheet.Charts.Add(ChartType.Line, 15, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];
            // Set the data source for the chart
            chart.NSeries.Add("C2:C5", true);          // Values
            chart.NSeries.CategoryData = "B2:B5";     // Dates as categories
            // Apply the same custom date format to X‑axis values
            chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";
            chart.Title.Text = "Sales Over Time";

            // ---------- Export the chart to a PDF document ----------
            chart.ToPdf("SalesTimelineChart.pdf");

            // Save the workbook (optional, to keep the timeline in the file)
            workbook.Save("SalesTimelineWorkbook.xlsx");
        }
    }
}
