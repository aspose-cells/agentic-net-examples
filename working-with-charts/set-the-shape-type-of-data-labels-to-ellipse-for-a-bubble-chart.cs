// Title: Aspose.Cells .NET: Set Bubble Chart Data Labels to Ellipse Shape (C#)
// Description: This example creates a workbook, fills X, Y, and size columns, adds a Bubble chart, binds a series, enables data labels, and changes the label shape to an ellipse using DataLabelShapeType.Ellipse. The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | .NET | bubble chart | data labels | ellipse shape | DataLabelShapeType | chart customization | Excel workbook | chart series
// Common Searches: Aspose.Cells set bubble chart data label shape ellipse | C# change data label shape to ellipse in Aspose.Cells | How to use DataLabelShapeType.Ellipse with Aspose.Cells | Enable and style bubble chart data labels Aspose.Cells | Aspose.Cells chart series data label shape options
// Developer Intent: Display bubble chart data labels as ellipses using Aspose.Cells for .NET.
// Use Cases: Generate a bubble chart where each point’s value appears inside an elliptical label. | Customize the visual appearance of a chart series by applying an ellipse shape to its data labels. | Create an Excel report that requires distinct, rounded data label styling for bubble charts.
// AI Prompts: Write C# code with Aspose.Cells that adds a bubble chart and sets its data labels to an ellipse shape. | Explain how DataLabelShapeType.Ellipse affects bubble chart data labels and show how to apply it in Aspose.Cells. | Provide step‑by‑step instructions to enable data labels and change their shape to ellipse for a chart series using Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsBubbleChartDataLabelEllipse
{
    // This example creates a workbook, fills X, Y, and size columns, adds a Bubble chart, binds a series, enables data labels, and changes the label shape to an ellipse using DataLabelShapeType.Ellipse. The workbook is then saved as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a bubble chart
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Size");
            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["C3"].PutValue(10);
            sheet.Cells["C4"].PutValue(15);

            // Add a bubble chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series and bind X, Y and bubble size ranges
            int seriesIndex = chart.NSeries.Add("B2:B4", true); // Y values
            chart.NSeries.CategoryData = "A2:A4";               // X values
            Series series = chart.NSeries[seriesIndex];
            series.BubbleSizes = "C2:C4";

            // Enable data labels and set their shape type to ellipse
            DataLabels dataLabels = series.DataLabels;
            dataLabels.ShowValue = true;                         // optional: show values
            dataLabels.ShapeType = DataLabelShapeType.Ellipse;   // set shape to ellipse

            // Save the workbook (lifecycle: save)
            workbook.Save("BubbleChartDataLabelEllipse.xlsx");
        }
    }
}
