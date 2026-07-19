// Title: Aspose.Cells .NET: Waterfall Chart with Custom Start, Intermediate, and Total Colors
// Description: Creates a workbook, populates categories and values, adds a Waterfall chart, then colors the start point (green), intermediate points (blue) and total point (red) before saving the Excel file.
// Keywords: Aspose.Cells | C# | .NET | Waterfall chart | custom point colors | ChartPoint foreground color | start point color | intermediate point color | total point color | Excel chart formatting | example code | GitHub | Aspose.Cells tutorial
// Common Searches: Aspose.Cells change waterfall chart point color C# | set start point color in Aspose.Cells waterfall chart | color intermediate bars in Aspose.Cells waterfall chart | apply total point color Aspose.Cells .NET | waterfall chart custom colors Aspose.Cells example
// Developer Intent: Apply distinct foreground colors to the start, middle, and total bars of a Waterfall chart using Aspose.Cells for .NET.
// Use Cases: Highlight opening balance in a financial waterfall chart with a green bar. | Visually separate quarterly changes from the final total using blue and red bars. | Generate a branded Excel report where key chart points follow corporate color guidelines.
// AI Prompts: Provide C# code that sets a specific ChartPoint foreground color in an Aspose.Cells Waterfall chart. | Show how to loop through Waterfall chart points and assign LightGreen to the first, LightBlue to the middle, and OrangeRed to the last point. | Explain the steps to bind data to a Waterfall chart and customize point colors based on position with Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, populates categories and values, adds a Waterfall chart, then colors the start point (green), intermediate points (blue) and total point (red) before saving the Excel file.
class WaterfallChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare data for the waterfall chart
        // Column A – categories, Column B – values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        string[] categories = { "Start", "Q1", "Q2", "Q3", "Total" };
        double[] values = { 100, 30, -20, 50, 160 }; // sample data

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(categories[i]); // A2, A3, ...
            sheet.Cells[i + 2, 1].PutValue(values[i]);   // B2, B3, ...
        }

        // Add a Waterfall chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the chart to the data range
        chart.NSeries.Add("B2:B6", true);          // values
        chart.NSeries.CategoryData = "A2:A6";      // categories

        // ----- Apply distinct colors -----
        // 1. Start point (first point)
        ChartPoint startPoint = chart.NSeries[0].Points[0];
        startPoint.Area.ForegroundColor = Color.LightGreen;

        // 2. Intermediate points (all points except first and last)
        for (int i = 1; i < chart.NSeries[0].Points.Count - 1; i++)
        {
            ChartPoint pt = chart.NSeries[0].Points[i];
            pt.Area.ForegroundColor = Color.LightBlue;
        }

        // 3. Total point (last point)
        ChartPoint totalPoint = chart.NSeries[0].Points[chart.NSeries[0].Points.Count - 1];
        totalPoint.Area.ForegroundColor = Color.OrangeRed;

        // Save the workbook to a file
        workbook.Save("WaterfallChartDemo.xlsx");
    }
}
