// Title: Set a fixed Z‑axis range of 0 to 100 for a 3‑D column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a workbook, add sample data, create a 3‑D column chart, and set its Z‑axis minimum to 0 and maximum to 100 with Aspose.Cells in C#. | Write C# code that builds a 3‑D column chart from a data range and configures the value axis to a static 0‑100 scale using Aspose.Cells.
// Common Searches: Aspose.Cells C# set Z axis minimum value for 3D column chart | How to fix the value axis range of a 3D chart to 0‑100 in Aspose.Cells | C# example for disabling automatic axis scaling in Aspose.Cells chart | Standardize Z‑axis scaling for 3‑D column chart with Aspose.Cells .NET
// Tags: Aspose.Cells set Z axis range | 3D column chart value axis scaling C# | disable automatic axis limits Aspose.Cells | chart value axis min max .NET | Aspose.Cells chart scaling example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart, ChartType, Axis

namespace AsposeCellsExamples
{
    // The example creates a new workbook, inserts sample data, adds a 3‑D column chart, disables automatic scaling, and explicitly sets the chart's Z‑axis (value axis) minimum to 0 and maximum to 100 before saving the file as ChartWithStandardizedZAxis.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(55);
                worksheet.Cells["B4"].PutValue(90);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the value (Z) axis scaling
                Axis valueAxis = chart.ValueAxis;
                valueAxis.IsAutomaticMinValue = false; // Disable automatic minimum
                valueAxis.IsAutomaticMaxValue = false; // Disable automatic maximum
                valueAxis.MinValue = 0;                // Set minimum to 0
                valueAxis.MaxValue = 100;              // Set maximum to 100

                // Define output file path
                string outputPath = "ChartWithStandardizedZAxis.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
