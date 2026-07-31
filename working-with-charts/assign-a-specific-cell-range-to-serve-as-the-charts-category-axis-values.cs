// Title: Set Category Axis Range for a Column Chart using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills columns A and B with labels and numbers, adds a column chart, assigns the Y‑axis series with NSeries.Add, and binds the X‑axis (category) labels to a specific cell range via the NSeries.CategoryData property before saving the file.
// Keywords: Aspose.Cells chart category axis | C# set X axis range | NSeries CategoryData property | column chart data binding Aspose.Cells | programmatic chart axis range .NET
// Common Searches: Aspose.Cells set chart X axis range C# | How to bind category labels to a chart in Aspose.Cells | Assign cell range to chart axis Aspose.Cells .NET | Chart NSeries CategoryData example
// Developer Intent: Bind a worksheet range to the chart’s category (X) axis.
// Use Cases: Generate a column chart where labels come from cells A2:A10. | Reuse the same category range across multiple series for consistent labeling. | Automatically update axis labels when the source table expands.
// AI Prompts: Show C# code to set a dynamic CategoryData range for an Aspose.Cells chart. | Explain how to change the category axis range after adding new rows to the data table. | Provide an example of linking a named range to a chart’s X‑axis using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryAxisDemo
{
    // Creates a workbook, fills columns A and B with labels and numbers, adds a column chart, assigns the Y‑axis series with NSeries.Add, and binds the X‑axis (category) labels to a specific cell range via the NSeries.CategoryData property before saving the file.
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
                sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series values (Y‑axis data)
            chart.NSeries.Add("=Sheet1!$B$2:$B$10", true);

            // Assign a specific cell range as the category (X‑axis) values
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$10";

            // Save the workbook
            workbook.Save("CategoryAxisDemo.xlsx");
        }
    }
}
