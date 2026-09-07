// Title: Generate an Excel audit worksheet that lists each chart’s axis type and tick‑label direction using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through every worksheet, extracts the primary Category, Value, and Series axes of each chart, and writes the axis type together with its tick‑label direction to a new worksheet named AuditReport. | Update the chart‑audit example to also capture each axis’s tick‑label font name and add that information to the audit worksheet. | Convert the audit program so that it writes the collected axis type and tick‑label direction data to a CSV file instead of an .xlsx workbook.
// Common Searches: aspocells how to list chart axis types and tick label direction in a workbook | c# iterate over charts in Excel file and get axis tick label orientation using Aspose.Cells | generate audit report of chart properties axis tick labels with Aspose.Cells .NET | retrieve chart axis tick label direction for each chart in Aspose.Cells workbook | export chart axis information to a separate worksheet using Aspose.Cells
// Tags: Aspose.Cells enumerate chart axes | Aspose.Cells retrieve tick label direction | Aspose.Cells write chart audit worksheet | Aspose.Cells export chart properties to CSV | Aspose.Cells auto‑fit columns after audit data

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAudit
{
    // The program creates a workbook, adds sample data and two charts, then scans every worksheet (excluding the audit sheet) to record for each chart the axis type (Category, Value, Series) and the tick‑label direction into an "AuditReport" worksheet, auto‑fits the columns, and saves the file as AuditReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data and charts (for demonstration only)
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx1 = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart1 = dataSheet.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B4", true);
            chart1.NSeries.CategoryData = "A2:A4";

            // Add a line chart
            int chartIdx2 = dataSheet.Charts.Add(ChartType.Line, 25, 0, 40, 10);
            Chart chart2 = dataSheet.Charts[chartIdx2];
            chart2.NSeries.Add("B2:B4", true);
            chart2.NSeries.CategoryData = "A2:A4";

            // Ensure charts are calculated before accessing axis properties
            chart1.Calculate();
            chart2.Calculate();

            // -------------------------------------------------
            // Create a worksheet to hold the audit report
            // -------------------------------------------------
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "AuditReport";

            // Header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("Chart Index");
            reportSheet.Cells[0, 2].PutValue("Axis Type");
            reportSheet.Cells[0, 3].PutValue("Tick Label Direction");

            int reportRow = 1; // start after header

            // -------------------------------------------------
            // Iterate through all worksheets and their charts
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the report sheet itself
                if (ws.Name == "AuditReport") continue;

                for (int i = 0; i < ws.Charts.Count; i++)
                {
                    Chart chart = ws.Charts[i];

                    // Category Axis (primary)
                    if (chart.HasAxis(AxisType.Category, true))
                    {
                        Axis catAxis = chart.CategoryAxis;
                        ChartTextDirectionType dir = catAxis.TickLabels.DirectionType;
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(i);
                        reportSheet.Cells[reportRow, 2].PutValue("Category");
                        reportSheet.Cells[reportRow, 3].PutValue(dir.ToString());
                        reportRow++;
                    }

                    // Value Axis (primary)
                    if (chart.HasAxis(AxisType.Value, true))
                    {
                        Axis valAxis = chart.ValueAxis;
                        ChartTextDirectionType dir = valAxis.TickLabels.DirectionType;
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(i);
                        reportSheet.Cells[reportRow, 2].PutValue("Value");
                        reportSheet.Cells[reportRow, 3].PutValue(dir.ToString());
                        reportRow++;
                    }

                    // Series Axis (if present)
                    if (chart.HasAxis(AxisType.Series, true))
                    {
                        Axis serAxis = chart.SeriesAxis;
                        ChartTextDirectionType dir = serAxis.TickLabels.DirectionType;
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(i);
                        reportSheet.Cells[reportRow, 2].PutValue("Series");
                        reportSheet.Cells[reportRow, 3].PutValue(dir.ToString());
                        reportRow++;
                    }
                }
            }

            // Auto-fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the workbook with the audit report
            workbook.Save("AuditReport.xlsx");
        }
    }
}
