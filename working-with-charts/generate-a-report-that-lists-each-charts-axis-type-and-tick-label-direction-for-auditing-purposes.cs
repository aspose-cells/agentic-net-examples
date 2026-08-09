// Title: Aspose.Cells for .NET – Generate a Chart Axis Audit Report (Axis Type & Tick‑Label Direction)
// Description: This C# example creates a workbook, adds sample data, inserts a column chart and a line chart, assigns different ChartTextDirectionType values to the Category and Value axes, forces chart calculation, and then iterates through every chart in the worksheet. For each chart it prints the axis type (Category, Value, and optional Series) together with the TickLabels.DirectionType, and finally saves the workbook. Ideal for developers who need to verify or document axis label orientation across multiple charts.
// Keywords: Aspose.Cells chart axis audit | C# chart tick label direction | ChartTextDirectionType .NET | list chart axis types Aspose | retrieve chart axis properties | Aspose.Cells console report | Excel chart axis enumeration | Aspose.Cells GitHub example | US developers Aspose.Cells | UK Aspose.Cells chart API
// Common Searches: How to list axis type and tick label direction for each chart using Aspose.Cells | Aspose.Cells C# generate chart axis audit report | Retrieve category and value axis label orientation in a workbook | Check for series axis in Aspose.Cells charts | Aspose.Cells chart axis properties example GitHub
// Developer Intent: Produce a console‑based audit that enumerates every chart’s axis type and its tick‑label direction.
// Use Cases: Document the label orientation of multiple charts before publishing an Excel report. | Validate that chart axes conform to corporate style guidelines (e.g., vertical category labels). | Detect and log the presence of a Series axis in charts that support it. | Automate quality checks for generated workbooks in CI pipelines.
// AI Prompts: Write C# code with Aspose.Cells that iterates through all charts in a worksheet and prints each axis’s type and TickLabels.DirectionType. | Create a method that returns a dictionary mapping chart names to a list of their axis types and corresponding tick‑label directions. | Show how to change the tick label direction for every value axis in a workbook and then generate an audit report.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, adds sample data, inserts a column chart and a line chart, assigns different ChartTextDirectionType values to the Category and Value axes, forces chart calculation, and then iterates through every chart in the worksheet. For each chart it prints the axis type (Category, Value, and optional Series) together with the TickLabels.DirectionType, and finally saves the workbook. Ideal for developers who need to verify or document axis label orientation across multiple charts.
class ChartAxisAudit
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for charts
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Value1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Value2");
        worksheet.Cells["C2"].PutValue(15);
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["C4"].PutValue(35);

        // Add first chart (Column)
        int chartIndex1 = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart1 = worksheet.Charts[chartIndex1];
        chart1.NSeries.Add("B2:B4", true);
        chart1.NSeries.CategoryData = "A2:A4";
        chart1.Title.Text = "Column Chart";
        // Set tick label directions for this chart
        chart1.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Vertical;
        chart1.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

        // Add second chart (Line)
        int chartIndex2 = worksheet.Charts.Add(ChartType.Line, 20, 0, 30, 10);
        Chart chart2 = worksheet.Charts[chartIndex2];
        chart2.NSeries.Add("C2:C4", true);
        chart2.NSeries.CategoryData = "A2:A4";
        chart2.Title.Text = "Line Chart";
        // Set tick label directions for this chart
        chart2.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;
        chart2.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate270;

        // Ensure axes are calculated before reading properties
        foreach (Chart ch in worksheet.Charts)
        {
            ch.Calculate();
        }

        // Generate audit report to console
        Console.WriteLine("Chart Axis Audit Report");
        Console.WriteLine("-----------------------");
        foreach (Chart ch in worksheet.Charts)
        {
            Console.WriteLine($"Chart Name: {ch.Name}");

            // Category axis
            Axis categoryAxis = ch.CategoryAxis;
            Console.WriteLine($"  Axis Type: Category");
            Console.WriteLine($"  Tick Label Direction: {categoryAxis.TickLabels.DirectionType}");

            // Value axis
            Axis valueAxis = ch.ValueAxis;
            Console.WriteLine($"  Axis Type: Value");
            Console.WriteLine($"  Tick Label Direction: {valueAxis.TickLabels.DirectionType}");

            // Series axis (if present)
            if (ch.HasAxis(AxisType.Series, true))
            {
                Axis seriesAxis = ch.SeriesAxis;
                Console.WriteLine($"  Axis Type: Series");
                Console.WriteLine($"  Tick Label Direction: {seriesAxis.TickLabels.DirectionType}");
            }
        }

        // Save the workbook containing the charts
        workbook.Save("ChartAxisAuditReport.xlsx");
    }
}
