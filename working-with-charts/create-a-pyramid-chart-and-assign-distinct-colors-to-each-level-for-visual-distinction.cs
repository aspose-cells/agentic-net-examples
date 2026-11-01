// Title: Create a pyramid chart with distinct colors for each level using Aspose.Cells for .NET
// AI Prompts: Generate a pyramid chart in a new workbook, bind it to three data rows, and set each pyramid level to a different color with Aspose.Cells. | Programmatically add level data, insert a Pyramid chart, and apply red, green, and blue colors to each chart point using Aspose.Cells for .NET.
// Common Searches: how to assign individual colors to pyramid chart points in Aspose.Cells C# | Aspose.Cells example for creating a pyramid chart with custom level colors | C# code to color each level of a pyramid chart differently in Excel | set foreground color for chart series points Aspose.Cells .NET | save pyramid chart with colored levels as XLSX using Aspose.Cells
// Tags: pyramid chart color customization Aspose.Cells | set point foreground color Aspose.Cells C# | add pyramid chart to worksheet Aspose.Cells | assign distinct colors to chart series points .NET | save pyramid chart as XLSX Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates an Excel workbook, populates three rows of level data, inserts a Pyramid chart, and assigns red, green, and blue colors to each chart point before saving the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells collection
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Populate data for the pyramid chart
            cells["A1"].PutValue("Level");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Level 1");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Level 2");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Level 3");
            cells["B4"].PutValue(10);

            // Add a pyramid chart to the worksheet (returns the chart index)
            int chartIndex = workbook.Worksheets[0].Charts.Add(ChartType.Pyramid, 5, 0, 20, 7);
            Chart chart = workbook.Worksheets[0].Charts[chartIndex];

            // Set the series data range and category labels
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Assign distinct colors to each level (point) in the pyramid
            Color[] levelColors = { Color.Red, Color.Green, Color.Blue };
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                chart.NSeries[0].Points[i].Area.ForegroundColor = levelColors[i % levelColors.Length];
            }

            // Determine output path and ensure the directory exists
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "PyramidChart.xlsx");
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
