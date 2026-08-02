// Title: Aspose.Cells .NET – Retrieve a Chart’s Parent Worksheet (Chart.Worksheet) and Log Its Name
// Description: C# example that creates a workbook, adds data, inserts a column chart, accesses the chart’s parent worksheet via the Chart.Worksheet property, writes the worksheet name to the console for diagnostics, and saves the file.
// Keywords: Aspose.Cells Chart.Worksheet | get chart parent worksheet .NET | log worksheet name from chart | C# Aspose.Cells chart worksheet reference | diagnostic logging Aspose.Cells chart
// Common Searches: How to get the worksheet that contains a chart using Aspose.Cells for .NET | Chart.Worksheet property example in C# | Retrieve chart parent sheet name Aspose.Cells | Aspose.Cells log chart worksheet name
// Developer Intent: Identify the worksheet a chart belongs to and output its name for troubleshooting or further processing.
// Use Cases: Confirm that a newly added chart is placed on the correct worksheet before publishing the workbook. | Batch‑process workbooks and record each chart’s parent sheet to aid error tracking. | Use the parent worksheet reference to modify related cells, apply formatting, or set tab colors based on chart characteristics.
// AI Prompts: Generate C# code that loops through all charts in a workbook and prints each chart’s parent worksheet name using Aspose.Cells. | Show how to change the tab color of a chart’s parent worksheet based on the chart type with Aspose.Cells for .NET. | Explain best practices for handling exceptions when accessing Chart.Worksheet in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds data, inserts a column chart, accesses the chart’s parent worksheet via the Chart.Worksheet property, writes the worksheet name to the console for diagnostics, and saves the file.
    public class ChartParentWorksheetDemo
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

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the chart's parent worksheet via Chart.Worksheet
                Worksheet parentWorksheet = chart.Worksheet;

                // Log the worksheet's name for diagnostic tracking
                Console.WriteLine("Chart's parent worksheet name: " + parentWorksheet.Name);

                // Determine output path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartParentWorksheetDemo_out.xlsx");

                // Save the workbook
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
            ChartParentWorksheetDemo.Run();
        }
    }
}
