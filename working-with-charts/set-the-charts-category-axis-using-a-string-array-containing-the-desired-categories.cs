// Title: Aspose.Cells .NET – Set Chart Category Axis Using a String Array
// Description: Learn how to assign custom category labels to a column chart in C# by setting the NSeries.CategoryData property with a brace‑enclosed, comma‑separated string array. The example creates a workbook, adds numeric series data, applies the string array as the category axis, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart CategoryData | string array categories | column chart axis labels | .NET Excel chart example | custom chart categories | Aspose.Cells CategoryData syntax | Excel export C#
// Common Searches: Aspose.Cells set chart categories from string array | CategoryData property C# example | how to change chart axis labels Aspose.Cells | assign custom category names to Excel chart .NET | string array for chart categories Aspose
// Developer Intent: Apply a predefined list of text labels to a chart's category axis via the CategoryData property.
// Use Cases: Build a quarterly sales column chart where the quarters (Q1‑Q4) are supplied as a string array. | Create a product comparison chart with product names defined in code and displayed on the axis. | Generate a regional performance chart by loading region names into a string array and setting them as categories before exporting.
// AI Prompts: Provide a C# snippet that sets chart CategoryData using a string array in Aspose.Cells. | Explain the exact brace‑enclosed syntax required for CategoryData when assigning custom axis labels. | Walk through the steps to create a column chart, add numeric series, apply string‑array categories, and save the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryAxisExample
{
    // Learn how to assign custom category labels to a column chart in C# by setting the NSeries.CategoryData property with a brace‑enclosed, comma‑separated string array. The example creates a workbook, adds numeric series data, applies the string array as the category axis, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some dummy numeric data for the series (values are required)
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"B{i}"].PutValue(i * 10); // 20,30,40,50
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data range (values only)
            chart.NSeries.Add("B2:B5", true);

            // Define categories using a string array syntax accepted by Aspose.Cells
            // The format is a comma‑separated list enclosed in braces.
            // Example: {"Jan","Feb","Mar","Apr"}
            chart.NSeries.CategoryData = "{\"Cat1\",\"Cat2\",\"Cat3\",\"Cat4\"}";

            // Optional: verify the CategoryData property
            Console.WriteLine("CategoryData set to: " + chart.NSeries.CategoryData);

            // Save the workbook
            workbook.Save("ChartWithStringArrayCategories.xlsx");
        }
    }
}
