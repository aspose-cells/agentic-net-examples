// Title: How to set a fixed vertical axis range (0 to 200) for a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart in a new workbook and forces the value axis to display from 0 to 200 using Aspose.Cells. | Show how to disable automatic scaling and assign custom MinValue and MaxValue to the vertical axis of an Excel chart with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells set chart value axis minimum to 0 and maximum to 200 in C# | How to turn off automatic axis scaling for a column chart with Aspose.Cells .NET | Programmatically define vertical axis limits for Excel charts using Aspose.Cells | C# example for fixing the Y‑axis range of an Aspose.Cells chart | Set fixed Y axis range 0‑200 for column chart in Aspose.Cells workbook
// Tags: Aspose.Cells custom vertical axis range | C# column chart fixed value axis | disable automatic chart scaling Aspose | Excel chart set MinValue MaxValue .NET | Aspose.Cells chart axis configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds sample data, inserts a column chart, disables automatic scaling, sets the vertical axis minimum to 0 and maximum to 200, and saves the file as ChartWithFixedVerticalAxis.xlsx.
    public class SetVerticalAxisScale
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(120);
                worksheet.Cells["B4"].PutValue(190);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the vertical (value) axis scaling
                Axis valueAxis = chart.ValueAxis;
                valueAxis.IsAutomaticMinValue = false; // Disable automatic minimum
                valueAxis.IsAutomaticMaxValue = false; // Disable automatic maximum
                valueAxis.MinValue = 0;                 // Set minimum to 0
                valueAxis.MaxValue = 200;               // Set maximum to 200

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithFixedVerticalAxis.xlsx");

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetVerticalAxisScale.Run();
        }
    }
}
