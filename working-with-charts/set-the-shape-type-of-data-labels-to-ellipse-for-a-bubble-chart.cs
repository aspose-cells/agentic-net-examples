// Title: Aspose.Cells C# – Set Bubble Chart Data Labels to Ellipse Shape
// Description: Creates a workbook, adds X/Y/size data, inserts a bubble chart, enables data labels, and changes the label shape to an ellipse using DataLabelShapeType.Ellipse before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | bubble chart | data labels | ellipse shape | DataLabelShapeType | chart formatting | Excel export
// Common Searches: Aspose.Cells set data label shape ellipse | C# bubble chart data labels Aspose.Cells | DataLabelShapeType.Ellipse example | change chart label shape Aspose.Cells
// Developer Intent: Apply an ellipse shape to data labels of a bubble chart.
// Use Cases: Display bubble values inside elliptical labels for clearer visualization. | Standardize label appearance across multiple series in a single chart. | Generate Excel reports where bubble chart labels use a custom shape.
// AI Prompts: Generate C# code with Aspose.Cells that creates a bubble chart and sets DataLabels.ShapeType to DataLabelShapeType.Ellipse. | Show how to format bubble chart data labels as ellipses using Aspose.Cells for .NET. | Explain the impact of DataLabelShapeType on label rendering and demonstrate applying the ellipse option.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds X/Y/size data, inserts a bubble chart, enables data labels, and changes the label shape to an ellipse using DataLabelShapeType.Ellipse before saving the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a bubble chart (X, Y, Size)
        worksheet.Cells["A1"].PutValue("X");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["A4"].PutValue(3);

        worksheet.Cells["B1"].PutValue("Y");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Size");
        worksheet.Cells["C2"].PutValue(5);
        worksheet.Cells["C3"].PutValue(10);
        worksheet.Cells["C4"].PutValue(15);

        // Add a bubble chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add a series to the chart and set its X, Y, and bubble size ranges
        int seriesIndex = chart.NSeries.Add("B2:B4", true); // Y values
        chart.NSeries.CategoryData = "A2:A4";               // X values
        Series series = chart.NSeries[seriesIndex];
        series.BubbleSizes = "C2:C4";

        // Enable data labels for the series
        DataLabels dataLabels = series.DataLabels;
        dataLabels.ShowValue = true;

        // Set the shape type of the data labels to ellipse
        dataLabels.ShapeType = DataLabelShapeType.Ellipse;

        // Save the workbook with the configured bubble chart
        workbook.Save("BubbleChartEllipseDataLabels.xlsx");
    }
}
