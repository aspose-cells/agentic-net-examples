// Title: Display values and percentages in data labels of a stacked bar chart and resize chart and worksheet shapes using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a stacked bar chart from range A1:B5, enable data labels to show each cell's value and its percentage, and set the plot area size to 600 × 400 points with Aspose.Cells in C#. | Iterate over every Shape object on a worksheet and set its Width to 200 points and Height to 100 points using the Aspose.Cells drawing API. | Create the output directory if it does not exist and save the workbook to a specified file path with Aspose.Cells.
// Common Searches: Aspose.Cells C# stacked bar chart show value and percentage in data labels | Resize chart plot area dimensions with Aspose.Cells .NET | Set worksheet shape width and height using Aspose.Cells drawing API | Create workbook with sample data when input file is missing Aspose.Cells | Verify output folder before saving workbook Aspose.Cells C#
// Tags: stacked bar chart data labels value and percentage Aspose.Cells | resize chart plot area dimensions Aspose.Cells C# | modify worksheet shape dimensions Aspose.Cells drawing | conditional workbook creation with sample data Aspose.Cells | ensure output folder exists before saving workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an existing workbook or creates a new one with sample data, adds a stacked bar chart based on A1:B5, enables data labels to display both cell values and percentages, resizes the chart's plot area to 600 × 400 points, adjusts every worksheet shape to 200 × 100 points, ensures the output directory exists, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input file exists; otherwise create a new workbook.
            string inputPath = "input.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with sample data.
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B4"].PutValue(30);
                ws.Cells["A5"].PutValue("D");
                ws.Cells["B5"].PutValue(40);
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a stacked bar chart.
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add data series (range A1:B5, categories in the first column).
            chart.NSeries.Add("A1:B5", true);

            // Configure data labels for each series.
            foreach (Series series in chart.NSeries)
            {
                // Enable data labels and show value & percentage.
                series.DataLabels.ShowValue = true;
                series.DataLabels.ShowPercentage = true;
            }

            // Resize the chart's plot area.
            chart.PlotArea.Width = 600;   // Width in points.
            chart.PlotArea.Height = 400;  // Height in points.

            // Optionally resize all other shapes on the worksheet.
            foreach (Shape shape in sheet.Shapes)
            {
                shape.Width = 200;
                shape.Height = 100;
            }

            // Ensure the output directory exists.
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any errors.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
