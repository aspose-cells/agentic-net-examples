// Title: C# – Export Chart Shape Connection Points to CSV with Aspose.Cells
// Description: This example creates a workbook, adds a column chart, inserts a rectangle shape inside the chart, retrieves the shape's connection points using GetConnectionPoints(), and writes each point's index, X and Y coordinates to a CSV file. The workbook can also be saved for further inspection.
// Keywords: Aspose.Cells GetConnectionPoints | chart shape coordinates C# | export shape points to CSV | Aspose.Cells shape connection points example | C# Aspose.Cells chart shape CSV export | retrieve shape anchor points Aspose | Aspose.Cells shape geometry extraction
// Common Searches: How to get connection points of a chart shape using Aspose.Cells | Export shape connection points to CSV in C# | Aspose.Cells GetConnectionPoints chart example | Save chart shape coordinates as CSV file | C# Aspose.Cells retrieve rectangle shape points
// Developer Intent: Extract all connection points of a chart shape and write them to a CSV file.
// Use Cases: Generate a CSV report of shape anchor points for layout analysis. | Feed shape geometry data into downstream reporting or visualization tools. | Automate custom connector logic by programmatically accessing shape connection points.
// AI Prompts: Write C# code that reads the CSV of shape connection points and creates a scatter chart from the data using Aspose.Cells. | Show how to resize the rectangle shape, recalculate its connection points, and export the updated coordinates to a new CSV file. | Explain how to iterate over every shape in a chart, collect each shape's connection points, and generate separate CSV files for each shape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a column chart, inserts a rectangle shape inside the chart, retrieves the shape's connection points using GetConnectionPoints(), and writes each point's index, X and Y coordinates to a CSV file. The workbook can also be saved for further inspection.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Add a rectangle shape inside the chart (this shape has connection points)
        Shape shape = chart.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            1000,   // left position (in 1/4000 of chart width)
            1000,   // top position (in 1/4000 of chart height)
            2000,   // width (in 1/4000 of chart width)
            1000,   // height (in 1/4000 of chart height)
            0,      // rotation angle
            0);     // flip mode
        shape.Text = "Demo Shape";

        // Retrieve all connection points of the shape
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Export the connection points to a CSV file
        string csvFile = "ChartShapeConnectionPoints.csv";
        using (StreamWriter writer = new StreamWriter(csvFile))
        {
            // Header
            writer.WriteLine("Index,X,Y");
            // Data rows
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                writer.WriteLine($"{i + 1},{connectionPoints[i][0]},{connectionPoints[i][1]}");
            }
        }

        // Save the workbook (optional, to keep the chart and shape)
        workbook.Save("ChartShapeDemo.xlsx");
    }
}
