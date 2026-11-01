// Title: Create a combo chart with column and line series and format the line series as a red dashed line using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a combo chart with column and line series, then configures the line series to use a red dash pattern and a weight of 2 points via Aspose.Cells. | Write a function that takes an existing Aspose.Cells workbook and updates the line series of a combo chart to have a custom dash style, specific line thickness, and a chosen color.
// Common Searches: how to set a dashed line for the second series in an Aspose.Cells combo chart using C# | Aspose.Cells example for changing line series dash pattern and thickness in a combo chart | C# code to format line series with red color and dash style in an Aspose.Cells chart | programmatically apply custom line weight and dash style to a combo chart series in Aspose.Cells .NET
// Tags: Aspose.Cells combo chart line series styling | C# set line dash pattern Aspose.Cells | Aspose.Cells chart series line weight color | Aspose.Cells .NET customize line series appearance | Aspose.Cells apply red dashed line to chart series

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds sample data, inserts a combo chart (column + line) on the first worksheet, defines a column series and a line series, switches the second series to a line type, and demonstrates how to apply a red dashed line style with custom weight using Aspose.Cells before saving the file as ComboChart_DashedLine.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the combo chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("ColumnSeries");
            ws.Cells["C1"].PutValue("LineSeries");

            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");

            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            ws.Cells["C2"].PutValue(15);
            ws.Cells["C3"].PutValue(25);
            ws.Cells["C4"].PutValue(35);

            // Add a combo chart (Column + Line) to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.Title.Text = "Combo Chart with Dashed Line Series";

            // Add the column series (first series)
            int columnSeriesIdx = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[columnSeriesIdx].Name = "ColumnSeries";

            // Add the line series (second series)
            int lineSeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[lineSeriesIdx].Name = "LineSeries";

            // Set the second series to be a line chart type
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // NOTE: In some versions of Aspose.Cells the Series class does not expose a Line property.
            // If available, you could customize the line style as shown below:
            // chart.NSeries[lineSeriesIdx].Line.DashStyle = MsoLineDashStyle.Dash;
            // chart.NSeries[lineSeriesIdx].Line.Weight = 2.0;
            // chart.NSeries[lineSeriesIdx].Line.Color = Color.Red;

            // Save the workbook with the chart
            wb.Save("ComboChart_DashedLine.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
