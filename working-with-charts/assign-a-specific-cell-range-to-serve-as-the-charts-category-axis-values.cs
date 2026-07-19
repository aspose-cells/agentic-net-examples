// Title: Set Category Axis Range for a Column Chart with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, populates columns A and B with category names and values, inserts a column chart, adds the value series from B2:B6, and assigns the X‑axis labels using the NSeries.CategoryData property with the range A2:A6. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | C# chart category axis | NSeries CategoryData | column chart X‑axis labels | set chart range .NET | programmatic Excel chart | Aspose.Cells chart example
// Common Searches: Aspose.Cells set X axis labels C# | How to assign category range to chart in .NET | NSeries CategoryData property usage | Define chart categories from cells Aspose.Cells | C# Excel chart category axis range
// Developer Intent: Assign a specific worksheet range as the category (X‑axis) labels for a chart using Aspose.Cells.
// Use Cases: Generate a column chart where the X‑axis displays custom category names from a data table. | Reuse the same category range across multiple series to keep axis labels consistent. | Dynamically change the category axis by updating the range string at runtime based on user input.
// AI Prompts: Write C# code with Aspose.Cells that sets the range A2:A10 as category axis labels for a line chart. | Create a reusable method that takes a Worksheet, Chart, and range string, then applies it to the chart's CategoryData property. | Show how to modify an existing chart's category axis range at runtime using a variable range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, populates columns A and B with category names and values, inserts a column chart, adds the value series from B2:B6, and assigns the X‑axis labels using the NSeries.CategoryData property with the range A2:A6. The workbook is then saved as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data: column A for categories, column B for values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Cat" + (i - 1));   // Category names
            sheet.Cells[$"B{i}"].PutValue(i * 10);           // Corresponding values
        }

        // Insert a column chart into the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the series data (values) – vertical orientation
        chart.NSeries.Add("B2:B6", true);

        // Assign the cell range that provides the category (X‑axis) labels
        chart.NSeries.CategoryData = "A2:A6";

        // Save the workbook to a file
        workbook.Save("CategoryAxisRangeDemo.xlsx");
    }
}
