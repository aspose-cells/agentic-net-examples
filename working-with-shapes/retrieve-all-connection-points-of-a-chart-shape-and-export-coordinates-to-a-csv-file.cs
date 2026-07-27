// Title: C# – Export Chart Shape Connection Points to CSV with Aspose.Cells
// Description: This example creates a workbook, adds a column chart and a free‑floating rectangle shape, extracts all connection points of the shape using GetConnectionPoints(), and writes the point index with X/Y coordinates to a CSV file. The workbook can also be saved for reference.
// Keywords: Aspose.Cells C# GetConnectionPoints | export shape coordinates CSV | .NET chart shape connection points | Aspose.Cells retrieve shape anchors | write shape points to file | chart helper shape Aspose.Cells
// Common Searches: Aspose.Cells get connection points of a shape | export shape anchor coordinates to CSV C# | how to retrieve chart shape connection points Aspose.Cells | C# example for GetConnectionPoints Aspose.Cells | save shape connection points as CSV
// Developer Intent: Extract every connection point of a chart‑related shape and save the coordinates to a CSV file.
// Use Cases: Debug layout of chart annotations by logging shape anchors. | Create a CSV report of shape connection points for external analytics. | Synchronize shape geometry with another system via exported coordinates.
// AI Prompts: Generate C# code that reads the ConnectionPoints.csv file and rebuilds the same rectangle shape in a new workbook using Aspose.Cells. | Show how to reposition or resize the rectangle shape based on its retrieved connection point values. | Explain the coordinate system returned by GetConnectionPoints and how it maps to worksheet rows and columns.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// This example creates a workbook, adds a column chart and a free‑floating rectangle shape, extracts all connection points of the shape using GetConnectionPoints(), and writes the point index with X/Y coordinates to a CSV file. The workbook can also be saved for reference.
class RetrieveChartShapeConnectionPoints
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a rectangle shape associated with the chart
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            Shape shape = sheet.Shapes.AddRectangle(5, 0, 0, 0, 100, 200);
            shape.Placement = PlacementType.FreeFloating; // set placement type
            shape.Text = "Chart Helper Shape";

            // Retrieve all connection points of the shape
            float[][] connectionPoints = shape.GetConnectionPoints();

            // Export the coordinates to a CSV file
            string csvPath = "ConnectionPoints.csv";
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                writer.WriteLine("PointIndex,X,Y");
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    float x = connectionPoints[i][0];
                    float y = connectionPoints[i][1];
                    writer.WriteLine($"{i + 1},{x},{y}");
                }
            }

            // Save the workbook (optional, to keep the chart and shape)
            string workbookPath = "ChartWithShape.xlsx";
            workbook.Save(workbookPath);

            Console.WriteLine($"Connection points exported to '{csvPath}'.");
            Console.WriteLine($"Workbook saved as '{workbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
