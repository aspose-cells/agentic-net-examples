// Title: Aspose.Cells .NET – Chart Axis Audit Report (Axis Type & Tick‑Label Direction)
// Description: C# program that builds a workbook, adds a column and a line chart (with secondary axes), calculates them, then scans all worksheets (except the report sheet) to record each chart’s name, axis type (Category, Value, Series), primary/secondary flag, and tick‑label direction into a new worksheet called "AxisAuditReport" and saves the file as ChartAxisAuditReport.xlsx.
// Keywords: Aspose.Cells | C# | .NET | chart axis audit | tick label direction | primary axis | secondary axis | ChartTextDirectionType | Excel automation | generate axis report
// Common Searches: Aspose.Cells list chart axes C# | How to get tick label direction of Excel chart with Aspose.Cells | Create axis audit worksheet using Aspose.Cells .NET | Enumerate primary and secondary axes in Excel charts | Chart axis type report Aspose.Cells
// Developer Intent: Generate a worksheet that enumerates every chart in a workbook, showing each axis’s type, whether it is primary or secondary, and its tick‑label direction.
// Use Cases: Audit chart formatting across a workbook to ensure consistent axis settings before publishing. | Validate that automatically generated reports have correctly configured secondary axes. | Document chart axis configurations for compliance reviews or technical documentation. | Create a reusable utility for developers to quickly inspect axis properties in large Excel files.
// AI Prompts: Write C# code with Aspose.Cells that iterates all charts in a workbook and writes each axis’s type and tick‑label direction to a new worksheet. | Refactor the axis‑audit example to extract the report‑writing logic into a separate reusable method. | Explain the possible values of ChartTextDirectionType and how to change them for localized tick‑label directions in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAxisAudit
{
    // C# program that builds a workbook, adds a column and a line chart (with secondary axes), calculates them, then scans all worksheets (except the report sheet) to record each chart’s name, axis type (Category, Value, Series), primary/secondary flag, and tick‑label direction into a new worksheet called "AxisAuditReport" and saves the file as ChartAxisAuditReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Sample data for demonstration (you can replace with your own data)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Value2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a primary chart (Column)
            int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart1 = sheet.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B4", true);
            chart1.NSeries.CategoryData = "A2:A4";
            chart1.Title.Text = "Primary Column Chart";

            // Add a secondary chart (Line) to demonstrate secondary axes
            int chartIdx2 = sheet.Charts.Add(ChartType.Line, 25, 0, 40, 12);
            Chart chart2 = sheet.Charts[chartIdx2];
            chart2.NSeries.Add("C2:C4", true);
            chart2.NSeries.CategoryData = "A2:A4";
            chart2.Title.Text = "Secondary Line Chart";

            // Enable secondary axes for the second chart
            chart2.HasAxis(AxisType.Value, false); // ensure secondary value axis exists
            chart2.HasAxis(AxisType.Category, false); // ensure secondary category axis exists

            // Calculate charts so that tick label information becomes available
            chart1.Calculate();
            chart2.Calculate();

            // -------------------------------------------------
            // Create a report worksheet
            // -------------------------------------------------
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "AxisAuditReport";

            // Write headers
            reportSheet.Cells["A1"].PutValue("Chart Name");
            reportSheet.Cells["B1"].PutValue("Axis Type");
            reportSheet.Cells["C1"].PutValue("Primary/Secondary");
            reportSheet.Cells["D1"].PutValue("Tick Label Direction");

            int reportRow = 1; // zero‑based index; row 1 is the second row (after headers)

            // -------------------------------------------------
            // Iterate through all worksheets and their charts
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the report sheet itself
                if (ws.Name == "AxisAuditReport") continue;

                foreach (Chart chart in ws.Charts)
                {
                    // Helper local function to write a line to the report
                    void WriteReport(string axisName, string primaryFlag, ChartTextDirectionType direction)
                    {
                        reportSheet.Cells[reportRow, 0].PutValue(chart.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(axisName);
                        reportSheet.Cells[reportRow, 2].PutValue(primaryFlag);
                        reportSheet.Cells[reportRow, 3].PutValue(direction.ToString());
                        reportRow++;
                    }

                    // Category Axis (primary)
                    if (chart.HasAxis(AxisType.Category, true))
                    {
                        Axis catAxis = chart.CategoryAxis;
                        ChartTextDirectionType dir = catAxis.TickLabels.DirectionType;
                        WriteReport("Category", "Primary", dir);
                    }

                    // Value Axis (primary)
                    if (chart.HasAxis(AxisType.Value, true))
                    {
                        Axis valAxis = chart.ValueAxis;
                        ChartTextDirectionType dir = valAxis.TickLabels.DirectionType;
                        WriteReport("Value", "Primary", dir);
                    }

                    // Series Axis (primary) – only for 3‑D charts
                    if (chart.HasAxis(AxisType.Series, true))
                    {
                        Axis serAxis = chart.SeriesAxis;
                        ChartTextDirectionType dir = serAxis.TickLabels.DirectionType;
                        WriteReport("Series", "Primary", dir);
                    }

                    // Category Axis (secondary)
                    if (chart.HasAxis(AxisType.Category, false))
                    {
                        Axis secCatAxis = chart.SecondCategoryAxis;
                        ChartTextDirectionType dir = secCatAxis.TickLabels.DirectionType;
                        WriteReport("Category", "Secondary", dir);
                    }

                    // Value Axis (secondary)
                    if (chart.HasAxis(AxisType.Value, false))
                    {
                        Axis secValAxis = chart.SecondValueAxis;
                        ChartTextDirectionType dir = secValAxis.TickLabels.DirectionType;
                        WriteReport("Value", "Secondary", dir);
                    }
                }
            }

            // -------------------------------------------------
            // Save the workbook with the report
            // -------------------------------------------------
            workbook.Save("ChartAxisAuditReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
