// Title: Aspose.Cells for .NET – attempting to apply plastic material and Soft lighting (70% intensity) to a 3‑D column chart (unsupported feature)
// Common Searches: how to set plastic material on a 3D chart using Aspose.Cells C# | apply soft lighting preset with 70% intensity to a 3D column chart in Aspose.Cells | does Aspose.Cells support material and lighting settings for 3D charts | Aspose.Cells 3D chart formatting options and limitations in .NET | set lighting intensity on a 3D chart with Aspose.Cells C#
// Tags: 3d chart material property Aspose.Cells | soft lighting preset Aspose.Cells | Aspose.Cells 3d chart formatting limitation | Aspose.Cells set chart material .NET | Aspose.Cells chart lighting intensity

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a 3‑D column chart with sample data, and saves it. It also documents that the current Aspose.Cells version does not support advanced 3‑D formatting such as assigning a plastic material or applying a Soft lighting preset with a specific intensity.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Get the first worksheet (default workbook already has one)
                Worksheet sheet = workbook.Worksheets[0];

                // Add a 3‑D column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Populate example data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add the data series to the chart
                chart.NSeries.Add("B2:B4", true);

                // NOTE: Advanced 3‑D formatting (material, lighting) is not available in the
                // current Aspose.Cells version, so those settings are omitted.

                // Define output file path
                string outputPath = "ThreeDChart.xlsx";

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
