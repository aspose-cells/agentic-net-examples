// Title: Create and configure a Pie chart in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a pie chart to an Excel sheet using Aspose.Cells in C# | example of setting a title for a pie chart with Aspose.Cells .NET | binding a data series to cells B2:B4 for a pie chart in Aspose.Cells | specify chart location by row and column indices in Aspose.Cells C# | save generated workbook as PieChart.xlsx with Aspose.Cells
// Tags: Aspose.Cells add pie chart | C# set pie chart title Aspose.Cells | Aspose.Cells bind series to range | Aspose.Cells chart positioning rows columns | Aspose.Cells save workbook as xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:B4 with fruit categories and values, adds a Pie chart (ChartType.Pie) positioned between rows 5‑20 and columns 0‑10, sets the chart title to "Fruit Distribution", binds the value series to the range B2:B4, and saves the file as PieChart.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(50);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Fruit Distribution";

            // Define the data series (values). Category data line removed because
            // the Series class in the current Aspose.Cells version does not expose
            // a CategoryData property.
            chart.NSeries.Add("B2:B4", true);

            // Determine output path and save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PieChart.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
