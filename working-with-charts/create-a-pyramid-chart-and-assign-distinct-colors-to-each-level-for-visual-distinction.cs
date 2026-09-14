// Title: Generate a Pyramid chart with individual colors for each level using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add sample category and series data, insert a Pyramid chart, set the first series fill to red, the second to green, the third to blue, and save the file as an .xlsx using the Aspose.Cells C# API. | Apply solid fill colors to each series of a Pyramid chart in Aspose.Cells, then export the workbook with distinct level colors.
// Common Searches: C# Aspose.Cells how to color each series in a pyramid chart separately | set custom fill colors for pyramid chart levels using Aspose.Cells .NET | example of creating a pyramid chart with red, green, blue series in C# | Aspose.Cells pyramid chart distinct colors per level tutorial
// Tags: pyramid chart series fill Aspose.Cells C# | assign solid fill color to chart series .NET | Aspose.Cells create colored pyramid chart | export workbook with pyramid chart .xlsx

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace PyramidChartDemo
{
    // The example creates a new workbook, populates category and series data, adds a Pyramid chart, assigns red, green, and blue solid fill colors to the three series representing pyramid levels, and saves the workbook as PyramidChartDistinctColors.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                // Categories (levels of the pyramid)
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("Level 1");
                cells["A3"].PutValue("Level 2");
                cells["A4"].PutValue("Level 3");

                // Series data – each series will become a distinct color slice
                cells["B1"].PutValue("Series 1");
                cells["B2"].PutValue(30);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(10);

                cells["C1"].PutValue("Series 2");
                cells["C2"].PutValue(20);
                cells["C3"].PutValue(15);
                cells["C4"].PutValue(5);

                cells["D1"].PutValue("Series 3");
                cells["D2"].PutValue(10);
                cells["D3"].PutValue(5);
                cells["D4"].PutValue(2);

                // Add a Pyramid chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Pyramid, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (all series)
                chart.NSeries.Add("B2:D4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Assign distinct colors to each series (level)
                // For Pyramid charts each series represents a level, so set the series area color.
                Series series0 = chart.NSeries[0];
                series0.Area.FillFormat.SolidFill.Color = Color.Red;      // Series 1 – Red

                Series series1 = chart.NSeries[1];
                series1.Area.FillFormat.SolidFill.Color = Color.Green;    // Series 2 – Green

                Series series2 = chart.NSeries[2];
                series2.Area.FillFormat.SolidFill.Color = Color.Blue;     // Series 3 – Blue

                // Save the workbook with the pyramid chart
                string outputPath = "PyramidChartDistinctColors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
