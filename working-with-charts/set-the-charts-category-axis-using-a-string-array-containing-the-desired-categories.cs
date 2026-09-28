// Title: How to assign custom string categories to a column chart’s category axis using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and sets the CategoryData property from a string[] of labels. | Transform a string array into the Aspose.Cells category data syntax (e.g., {"Label1","Label2"}) and apply it to a chart’s category axis. | Demonstrate programmatic assignment of custom category labels to a chart series in a .NET workbook using Aspose.Cells.
// Common Searches: aspnet set chart category axis labels from string array aspose.cells | c# aspose.cells column chart custom category labels example | how to format CategoryData property with string array in Aspose.Cells | assigning custom categories to chart series programmatically using Aspose.Cells .NET
// Tags: Aspose.Cells set chart CategoryData from string array | C# column chart custom category axis Aspose.Cells | CategoryData syntax for string labels .NET | programmatic chart label assignment Aspose.Cells | Aspose.Cells chart series category customization

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryAxisExample
{
    // The example creates a new workbook, adds numeric data, inserts a column chart, converts a C# string array into the CategoryData format required by Aspose.Cells, assigns it to the chart's category axis, and saves the workbook as ChartWithStringCategories.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample numeric data for the chart series (required for the chart to have values)
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values only, categories will be set separately)
            chart.NSeries.Add("B2:B4", true);

            // Define the desired categories as a string array
            string[] categories = new string[] { "Alpha", "Beta", "Gamma" };

            // Convert the string array to the format accepted by CategoryData (e.g., {"Alpha","Beta","Gamma"})
            string categoryData = "{" + string.Join(",", categories.Select(c => $"\"{c}\"")) + "}";

            // Assign the category axis using the prepared string sequence
            chart.NSeries.CategoryData = categoryData;

            // Optional: verify the assigned CategoryData (output to console)
            Console.WriteLine("Assigned CategoryData: " + chart.NSeries.CategoryData);

            // Save the workbook to a file
            workbook.Save("ChartWithStringCategories.xlsx");
        }
    }
}
