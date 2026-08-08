// Title: Aspose.Cells for .NET – Position a Column Chart at Row 15, Column 3 and Set Width to 400 pt
// Description: Creates a workbook, adds sample data, inserts a column chart, moves the chart so its upper‑left corner starts at row 15 column 3 using Chart.Move, sets the chart width to 400 points via ChartObject.WidthPt, and saves the file as an XLSX document.
// Keywords: Aspose.Cells chart positioning | Chart.Move row column .NET | ChartObject.WidthPt | set chart width points | column chart placement Aspose.Cells | C# Aspose.Cells chart size | Excel chart layout programmatically
// Common Searches: Aspose.Cells move chart to specific row and column | set chart width to 400 points in C# | how to position a chart at row 15 column 3 using Aspose.Cells | adjust chart dimensions with ChartObject.WidthPt | C# example for chart placement in Excel workbook
// Developer Intent: Place a column chart at worksheet row 15, column 3 and define its width as 400 points.
// Use Cases: Generate a formatted report where the chart must align with a predefined grid. | Design a dashboard layout that requires exact chart width for visual consistency. | Re‑locate an existing chart without modifying its data series or type.
// AI Prompts: Write C# code with Aspose.Cells that moves a chart to row 15, column 3 and sets its width to 400 pt. | Explain how the parameters of Chart.Move correspond to worksheet rows and columns and how to calculate bottomRow/rightColumn for a target size. | Show how to specify chart dimensions in inches or centimeters instead of points using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPositionDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, moves the chart so its upper‑left corner starts at row 15 column 3 using Chart.Move, sets the chart width to 400 points via ChartObject.WidthPt, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart with an initial position (rows 5‑15, columns 0‑5)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Move the chart so that its upper‑left corner starts at row 15, column 3
            // (bottomRow and rightColumn are set to keep the chart size reasonable)
            chart.Move(15, 3, 25, 8);

            // Set the chart width to 400 points (1 point = 1/72 inch)
            chart.ChartObject.WidthPt = 400;

            // Optionally, you can also set the height if needed
            // chart.ChartObject.HeightPt = 300;

            // Save the workbook
            workbook.Save("ChartPositioned.xlsx", SaveFormat.Xlsx);
        }
    }
}
