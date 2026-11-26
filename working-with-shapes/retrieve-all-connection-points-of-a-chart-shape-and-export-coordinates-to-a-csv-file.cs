// Title: Extract chart shape connection points and write them to a CSV file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get chart shape connection points and export to CSV | How to retrieve coordinates of a chart object using Aspose.Cells .NET | Saving chart shape connection points to a CSV file in C# | Example of GetConnectionPoints method for Aspose.Cells chart shapes | Exporting Aspose.Cells chart connection points for further analysis
// Tags: chartshape getconnectionpoints aspnet | export chart shape coordinates csv aspnet | aspocells chart connection points extraction | write shape points to csv csharp | aspocells chartobject coordinate export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, obtains the ChartShape via chart.ChartObject, calls GetConnectionPoints() to retrieve all connection points, and writes each point's index, X and Y values to a CSV file before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and add sample data for the chart.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // -----------------------------------------------------------------
            // 2. Add a chart to the worksheet.
            // -----------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -----------------------------------------------------------------
            // 3. Obtain the ChartShape (the visual shape of the chart).
            // -----------------------------------------------------------------
            ChartShape chartShape = chart.ChartObject;

            // -----------------------------------------------------------------
            // 4. Retrieve all connection points of the chart shape.
            // -----------------------------------------------------------------
            float[][] connectionPoints = chartShape.GetConnectionPoints();

            if (connectionPoints == null || connectionPoints.Length == 0)
            {
                Console.WriteLine("No connection points were found for the chart shape.");
                return;
            }

            // -----------------------------------------------------------------
            // 5. Export the coordinates to a CSV file.
            // -----------------------------------------------------------------
            string csvPath = "ChartConnectionPoints.csv";
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                writer.WriteLine("Index,X,Y");
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    float x = connectionPoints[i][0];
                    float y = connectionPoints[i][1];
                    writer.WriteLine($"{i + 1},{x},{y}");
                }
            }

            Console.WriteLine($"Successfully exported {connectionPoints.Length} connection points to \"{csvPath}\".");

            // Optional: Save the workbook to verify the chart exists.
            workbook.Save("ChartWithConnectionPoints.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
