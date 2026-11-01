// Title: How to set the depth of a 3‑D Column chart to 15 points using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a Column3D chart in a workbook and assigns a depth of fifteen points via the chart's ThreeDFormat.Depth property using Aspose.Cells. | Update an existing Chart object in an Aspose.Cells workbook to enable three‑dimensional formatting and set its depth to fifteen points.
// Common Searches: Aspose.Cells C# three dimensional chart depth 15 points | How to adjust 3D column chart depth in Aspose.Cells .NET | C# example for Aspose.Cells ThreeDFormat.Depth property | Increase 3D chart depth in an Excel file with Aspose.Cells | Set chart shape depth points using Aspose.Cells for .NET
// Tags: Aspose.Cells ThreeDFormat.Depth | configure 3D chart depth C# | column3d chart depth property | enable three‑dimensional formatting Aspose.Cells | chart shape depth points .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, fills cells A1:B4 with sample data, adds a 3‑D column chart linked to that data, sets the chart’s three‑dimensional depth to 15 points, and saves the file as ThreeDimensionalChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a 3‑D column chart (position: rows 5‑15, columns 0‑5)
                int chartIdx = sheet.Charts.Add(ChartType.Column3D, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];

                // Define the data series including categories (A2:A4) and values (B2:B4)
                chart.NSeries.Add("A2:B4", true);

                // Determine output path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ThreeDimensionalChart.xlsx");

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
