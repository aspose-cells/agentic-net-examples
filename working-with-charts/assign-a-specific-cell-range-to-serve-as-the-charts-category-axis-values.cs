// Title: Assign a worksheet cell range as the X‑axis category labels for a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that links a worksheet range to a chart’s category axis using Aspose.Cells. | Show how to use Aspose.Cells to set X‑axis labels from cells A2:A10 for a column chart. | Provide an example of binding worksheet data to a chart’s category axis in C# with Aspose.Cells.
// Common Searches: how to set X axis labels from worksheet cells in Aspose.Cells C# | Aspose.Cells programmatic column chart category range example | C# code for assigning category axis range to Aspose.Cells chart | binding worksheet cells to chart categories in a .NET workbook
// Tags: Aspose.Cells NSeries.CategoryData binding | C# column chart X‑axis label assignment | chart category data range definition Aspose.Cells | programmatic chart data source configuration .NET | worksheet range to chart category mapping

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryDataExample
{
    // The example creates a workbook, fills column A with category labels and column B with values, adds a column chart, sets the series values to B2:B10, assigns the X‑axis (category) labels to A2:A10 via the chart’s NSeries.CategoryData property, and saves the file as CategoryDataDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series values (Y‑axis data)
            chart.NSeries.Add("=Sheet1!$B$2:$B$10", true);

            // Assign the cell range that will be used for the category (X‑axis) values
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$10";

            // Save the workbook
            workbook.Save("CategoryDataDemo.xlsx");
        }
    }
}
