// Title: Aspose.Cells C# – Change Bubble Chart Data Labels to Ellipse Shape
// Description: Demonstrates how to create a workbook, add a bubble chart, enable data labels for a series, assign DataLabelShapeType.Ellipse so each label appears as an oval, and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# | bubble chart | data label shape | ellipse | DataLabelShapeType | chart customization | Excel automation | Aspose.Cells examples
// Common Searches: Aspose.Cells set data label ellipse | C# bubble chart label shape | How to change chart data label shape in Aspose.Cells | DataLabelShapeType.Ellipse example | Customize bubble chart labels Aspose.Cells .NET
// Developer Intent: Apply an elliptical shape to the data labels of a bubble‑chart series using Aspose.Cells for .NET.
// Use Cases: Design dashboards where bubble‑chart labels need a distinct oval background. | Generate client‑facing Excel reports with branded ellipse‑shaped data labels. | Update existing spreadsheets programmatically to modify label appearance for better readability.
// AI Prompts: Generate C# code with Aspose.Cells that adds a bubble chart and sets its series data labels to an ellipse. | Show how to enable and style data labels on a bubble chart using DataLabelShapeType.Ellipse in Aspose.Cells. | Explain step‑by‑step how to customize the shape of chart data labels in a .NET Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a bubble chart, enable data labels for a series, assign DataLabelShapeType.Ellipse so each label appears as an oval, and save the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the bubble chart
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

        // Add a series to the chart and set X, Y, and bubble size ranges
        int seriesIndex = chart.NSeries.Add("B2:B4", true); // Y values
        chart.NSeries.CategoryData = "A2:A4";               // X values
        Series series = chart.NSeries[seriesIndex];
        series.BubbleSizes = "C2:C4";                      // Bubble sizes

        // Enable data labels and set their shape type to ellipse
        DataLabels dataLabels = series.DataLabels;
        dataLabels.ShowValue = true;                       // Show the Y value in the label
        dataLabels.ShapeType = DataLabelShapeType.Ellipse; // Set shape to ellipse

        // Save the workbook with the configured chart
        workbook.Save("BubbleChartDataLabelEllipse.xlsx");
    }
}
