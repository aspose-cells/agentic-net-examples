// Title: Remove a Specific Series from an Aspose.Cells Chart by Index (C#)
// Description: Creates a workbook, adds sample data, builds a column chart with three series, removes the second series using NSeries.RemoveAt(1), shows the series count before and after removal, and saves the file.
// Keywords: Aspose.Cells chart series removal | C# delete chart series by index | Aspose.Cells NSeries.RemoveAt example | remove column chart series Aspose.Cells | Aspose.Cells chart manipulation .NET | Aspose.Cells remove series C# | Aspose.Cells chart API | Aspose.Cells GitHub example
// Common Searches: how to delete a series from an Aspose.Cells chart C# | remove second series from column chart Aspose.Cells | Aspose.Cells NSeries.RemoveAt usage | delete specific chart series by index .NET | Aspose.Cells chart series remove tutorial
// Developer Intent: The developer needs to programmatically delete a particular data series from a chart using its zero‑based index.
// Use Cases: Clean up a multi‑series chart by removing unwanted series before exporting the workbook. | Allow end‑users to select which series to display, dynamically removing the others at runtime. | Generate concise reports that show only the primary data series, eliminating secondary series that add visual clutter.
// AI Prompts: Generate C# code with Aspose.Cells that removes the third series from a line chart and updates the chart title. | Show how to check that a series index exists before calling RemoveAt to avoid runtime errors. | Explain how to delete multiple series in a loop with Aspose.Cells, handling index shifts after each removal.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a column chart with three series, removes the second series using NSeries.RemoveAt(1), shows the series count before and after removal, and saves the file.
    public class RemoveSeriesByIndexDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");
                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);
                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(150);
                worksheet.Cells["C3"].PutValue(250);
                worksheet.Cells["C4"].PutValue(350);
                worksheet.Cells["D1"].PutValue("Series3");
                worksheet.Cells["D2"].PutValue(120);
                worksheet.Cells["D3"].PutValue(220);
                worksheet.Cells["D4"].PutValue(320);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Add three series to the chart
                chart.NSeries.Add("B2:B4", true); // Series1
                chart.NSeries.Add("C2:C4", true); // Series2
                chart.NSeries.Add("D2:D4", true); // Series3
                chart.NSeries.CategoryData = "A2:A4";

                // Display initial series count
                Console.WriteLine($"Initial series count: {chart.NSeries.Count}");

                // Remove the second series (index 1)
                chart.NSeries.RemoveAt(1);

                // Display series count after removal
                Console.WriteLine($"Series count after removal: {chart.NSeries.Count}");

                // Save the workbook with the modified chart
                string outputPath = "RemoveSeriesByIndexDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                RemoveSeriesByIndexDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
