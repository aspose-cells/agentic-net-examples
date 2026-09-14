// Title: Convert an existing Excel pie chart to a doughnut chart by setting the chart’s Type with Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx workbook, locate the first chart, and assign ChartType.Doughnut to its Type using Aspose.Cells in C#. | Programmatically replace a pie chart with a doughnut chart in an existing Excel file by updating the chart’s Type property via the Aspose.Cells API. | Create a sample workbook containing a pie chart, then demonstrate how to switch the chart to doughnut format and save the file with C#.
// Common Searches: Aspose.Cells C# change chart type from pie to doughnut | How to set chart Type to Doughnut in an existing Excel workbook using Aspose.Cells | Update first chart in .xlsx to doughnut chart programmatically .NET | Convert pie chart to doughnut chart Aspose.Cells example code
// Tags: Aspose.Cells chart property update | C# doughnut chart generation | Excel workbook chart modification | Aspose.Cells chart type conversion | pie chart to doughnut transformation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing workbook (or creates one with a pie chart), changes the first chart's Type to ChartType.Doughnut, and saves the modified workbook as a new .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "ExistingPieChart.xlsx";
            const string outputPath = "UpdatedDoughnutChart.xlsx";

            Workbook workbook;

            // Load existing workbook or create a sample one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = CreateSampleWorkbookWithPieChart();
                workbook.Save(inputPath);
                Console.WriteLine($"Input file not found. Created sample workbook '{inputPath}'.");
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Change the first chart's type to Doughnut if a chart exists
            if (sheet.Charts.Count > 0)
            {
                Chart chart = sheet.Charts[0];
                chart.Type = ChartType.Doughnut;
                Console.WriteLine("Chart type changed to Doughnut.");
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a workbook containing a simple pie chart for demonstration purposes
    private static Workbook CreateSampleWorkbookWithPieChart()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Sample data
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["B2"].PutValue(30);
        ws.Cells["A3"].PutValue("B");
        ws.Cells["B3"].PutValue(70);

        // Add a pie chart
        int chartIndex = ws.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        return wb;
    }
}
