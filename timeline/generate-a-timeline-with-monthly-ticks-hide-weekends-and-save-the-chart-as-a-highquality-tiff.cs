using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

class TimelineChartTiffDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate worksheet with date‑based data ----------
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Value";

            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();

            // Add 12 months of data (including weekends)
            for (int i = 0; i < 12; i++)
            {
                sheet.Cells[i + 1, 0].Value = startDate.AddMonths(i);
                sheet.Cells[i + 1, 1].Value = rnd.Next(50, 200);
            }

            // ---------- Create a PivotTable (required for Timeline) ----------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B13", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a Timeline linked to the PivotTable ----------
            // Placed at row 0, column 5 (F1 cell) and bound to the "Date" field
            sheet.Timelines.Add(pivot, 0, 5, "Date");

            // ---------- Add a Line chart based on the original data ----------
            int chartIdx = sheet.Charts.Add(ChartType.Line, 15, 0, 30, 15);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B13", true);          // Values
            chart.NSeries.CategoryData = "A2:A13";     // Dates

            // ---------- Configure the category axis for monthly ticks ----------
            Axis categoryAxis = chart.CategoryAxis;
            categoryAxis.CategoryType = CategoryType.TimeScale;   // Enable time scale
            categoryAxis.BaseUnitScale = TimeUnit.Months;        // Base unit = months
            categoryAxis.MajorUnitScale = TimeUnit.Months;       // Major ticks = months
            categoryAxis.MinorUnitScale = TimeUnit.Days;         // Minor ticks = days
            categoryAxis.MajorUnit = 1;                          // One month per major tick

            // The property IsDateAxis enables date axis handling (default is true for time scale)
            // No additional code needed here.

            // Save the workbook to verify output (optional)
            string outputPath = "TimelineChartDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}