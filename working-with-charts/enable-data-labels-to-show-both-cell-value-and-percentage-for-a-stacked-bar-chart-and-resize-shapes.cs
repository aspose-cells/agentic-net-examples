// Title: Display cell values and percentages in stacked bar chart data labels and resize chart and shapes with Aspose.Cells for .NET
// AI Prompts: Create a stacked bar chart, enable its data labels to show both the cell value and the percentage, and assign a fixed pixel width to each label using Aspose.Cells. | Resize the chart object to 800 × 400 pixels and insert a rectangle shape inside the chart area by specifying left, top, right, and bottom as percentages of the chart dimensions.
// Common Searches: Aspose.Cells show value and percentage in stacked bar chart data labels | set fixed pixel width for data labels in Aspose.Cells chart | resize chart shape to specific pixel dimensions using Aspose.Cells .NET | add rectangle shape to chart area with scale (percentage) coordinates Aspose.Cells | prevent automatic data label resizing in stacked bar chart Aspose.Cells
// Tags: stacked bar chart data labels value percentage Aspose.Cells | custom data label width pixels Aspose.Cells | chart object resize pixels Aspose.Cells .NET | insert rectangle shape using scale coordinates Aspose.Cells | disable data label auto‑resize Aspose.Cells

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, builds a stacked bar chart, configures each series to display both cell values and percentages in data labels with an 80‑pixel fixed width, resizes the chart to 800 × 400 pixels, and places a light‑blue rectangle inside the chart area using 10‑30% scale coordinates, then saves the file as StackedBarChartWithLabels.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a stacked bar chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["C1"].PutValue("Series2");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["C2"].PutValue(15);
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["C4"].PutValue(35);

        // Add a stacked bar chart (BarStacked)
        int chartIndex = worksheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:C4", true);          // values
        chart.NSeries.CategoryData = "A2:A4";      // categories

        // Enable data labels to show both value and percentage for each series
        foreach (Series series in chart.NSeries)
        {
            series.DataLabels.ShowValue = true;            // show cell value
            series.DataLabels.ShowPercentage = true;       // show percentage
            // Prevent automatic resizing so we can set a custom width
            series.DataLabels.IsResizeShapeToFitText = false;
            series.DataLabels.WidthPixel = 80;             // custom width in pixels
        }

        // Resize the whole chart shape (ChartObject is a ChartShape)
        ChartShape chartShape = chart.ChartObject;
        chartShape.Width = 800;   // width in pixels
        chartShape.Height = 400;  // height in pixels

        // Add a rectangle shape inside the chart area using scale coordinates (percentage of chart area)
        Shape rectangle = chart.Shapes.AddShapeInChartByScale(
            MsoDrawingType.Rectangle,
            PlacementType.Move,
            0.1,   // left 10% of chart width
            0.1,   // top 10% of chart height
            0.3,   // right 30% of chart width
            0.3);  // bottom 30% of chart height

        rectangle.Fill.SolidFill.Color = Color.LightBlue;
        rectangle.Line.SolidFill.Color = Color.DarkBlue;

        // Save the workbook
        workbook.Save("StackedBarChartWithLabels.xlsx");
    }
}
