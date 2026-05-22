using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAuditReport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a custom tick label direction for demonstration
            chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Vertical;
            chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

            // Calculate the chart to ensure axis information is up‑to‑date
            chart.Calculate();

            // Add a worksheet to hold the audit report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "AuditReport";

            // Write header row
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Chart Index");
            reportSheet.Cells["C1"].PutValue("Axis Type");
            reportSheet.Cells["D1"].PutValue("Primary/Secondary");
            reportSheet.Cells["E1"].PutValue("Tick Label Direction");

            int reportRow = 1; // zero‑based index; row 1 is the second row (after header)

            // Iterate through all worksheets (excluding the report sheet itself)
            for (int wsIdx = 0; wsIdx < workbook.Worksheets.Count; wsIdx++)
            {
                Worksheet ws = workbook.Worksheets[wsIdx];
                if (ws.Name == "AuditReport") continue;

                // Iterate through all charts in the worksheet
                for (int cIdx = 0; cIdx < ws.Charts.Count; cIdx++)
                {
                    Chart ch = ws.Charts[cIdx];

                    // Helper local function to record an axis entry
                    void RecordAxis(string axisName, bool isPrimary, Axis axisObj)
                    {
                        // Get the direction of tick labels; default to Horizontal if null
                        ChartTextDirectionType direction = ChartTextDirectionType.Horizontal;
                        if (axisObj != null && axisObj.TickLabels != null)
                            direction = axisObj.TickLabels.DirectionType;

                        // Write data to the report sheet
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(cIdx);
                        reportSheet.Cells[reportRow, 2].PutValue(axisName);
                        reportSheet.Cells[reportRow, 3].PutValue(isPrimary ? "Primary" : "Secondary");
                        reportSheet.Cells[reportRow, 4].PutValue(direction.ToString());
                        reportRow++;
                    }

                    // Category Axis (primary)
                    if (ch.HasAxis(AxisType.Category, true))
                        RecordAxis("Category", true, ch.CategoryAxis);

                    // Category Axis (secondary)
                    if (ch.HasAxis(AxisType.Category, false))
                        RecordAxis("Category", false, ch.SecondCategoryAxis);

                    // Value Axis (primary)
                    if (ch.HasAxis(AxisType.Value, true))
                        RecordAxis("Value", true, ch.ValueAxis);

                    // Value Axis (secondary)
                    if (ch.HasAxis(AxisType.Value, false))
                        RecordAxis("Value", false, ch.SecondValueAxis);

                    // Series Axis (primary) – always primary for series axis
                    if (ch.HasAxis(AxisType.Series, true))
                        RecordAxis("Series", true, ch.SeriesAxis);
                }
            }

            // Save the workbook with the audit report
            workbook.Save("ChartAxisAuditReport.xlsx");
        }
    }
}