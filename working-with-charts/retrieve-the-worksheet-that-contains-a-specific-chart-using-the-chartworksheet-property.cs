// Title: Get the Worksheet Containing a Chart via Chart.Worksheet in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use the Chart.Worksheet property in Aspose.Cells for .NET to retrieve the parent Worksheet of a chart. The sample creates a workbook, adds data and a column chart, then prints the chart's worksheet index and name before saving the file.
// Keywords: Aspose.Cells | Chart.Worksheet | C# | .NET | retrieve chart worksheet | parent worksheet of chart | Aspose.Cells chart example | access worksheet from chart | Aspose.Cells API | chart parent sheet
// Common Searches: Aspose.Cells Chart.Worksheet example | how to get worksheet of a chart in C# | retrieve parent sheet of Aspose chart | chart worksheet index Aspose.Cells | C# Aspose.Cells get chart's worksheet name
// Developer Intent: Find the Worksheet object that hosts a specific Chart using the Chart.Worksheet property.
// Use Cases: Determine which sheet a chart belongs to when working with workbooks that contain multiple worksheets. | Log or display the worksheet name and index for debugging or reporting purposes. | Perform further modifications on the chart's parent sheet after locating it.
// AI Prompts: Write a C# snippet that adds a column chart to a worksheet with Aspose.Cells and then prints the chart's parent worksheet name and index using Chart.Worksheet. | Generate code that iterates over all charts in a workbook and outputs each chart’s containing worksheet name and index. | Create a method that accepts a Chart object and safely returns its Worksheet, handling null references and providing meaningful error messages.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to use the Chart.Worksheet property in Aspose.Cells for .NET to retrieve the parent Worksheet of a chart. The sample creates a workbook, adds data and a column chart, then prints the chart's worksheet index and name before saving the file.
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

                // Add sample data for the chart
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

                // Set chart data range
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve the worksheet that contains this chart
                Worksheet chartParentWorksheet = chart.Worksheet;

                // Output worksheet information to verify
                Console.WriteLine("Chart's worksheet index: " + chartParentWorksheet.Index);
                Console.WriteLine("Chart's worksheet name: " + chartParentWorksheet.Name);

                // Save the workbook
                string outputPath = "RetrieveChartWorksheetDemo_out.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
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
