// Title: Get a Chart’s Parent Worksheet with Aspose.Cells Chart.Worksheet in C#
// Description: Creates a workbook, adds a column chart, and uses the Chart.Worksheet property to retrieve the worksheet that hosts the chart, then prints its index and name before saving the file.
// Keywords: Aspose.Cells | Chart.Worksheet | C# | .NET | retrieve chart worksheet | chart parent sheet | worksheet index | worksheet name | Aspose.Cells chart example
// Common Searches: Aspose.Cells get worksheet of chart | Chart.Worksheet property C# | how to find chart's sheet Aspose.Cells | retrieve chart parent worksheet .NET | Aspose.Cells chart location
// Developer Intent: Obtain the worksheet object that contains a specific chart via the Chart.Worksheet property.
// Use Cases: Determine which sheet a chart belongs to before modifying its data source. | Log chart location (sheet name and index) for debugging or audit trails. | Validate chart placement when generating multi‑sheet reports. | Move or copy a chart after identifying its source worksheet.
// AI Prompts: Write C# code with Aspose.Cells that prints the name and index of a chart’s worksheet using Chart.Worksheet. | Show how to loop through all charts in a workbook and output each chart’s worksheet name. | Explain how to relocate a chart to a different worksheet using the Chart.Worksheet property in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, and uses the Chart.Worksheet property to retrieve the worksheet that hosts the chart, then prints its index and name before saving the file.
    public class RetrieveChartWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve the worksheet that contains this chart using Chart.Worksheet property
                Worksheet chartWorksheet = chart.Worksheet;

                // Output worksheet information to verify the property works
                Console.WriteLine("Chart's worksheet index: " + chartWorksheet.Index);
                Console.WriteLine("Chart's worksheet name: " + chartWorksheet.Name);

                // Save the workbook
                string outputPath = "RetrieveChartWorksheetDemo_out.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveChartWorksheetDemo.Run();
        }
    }
}
