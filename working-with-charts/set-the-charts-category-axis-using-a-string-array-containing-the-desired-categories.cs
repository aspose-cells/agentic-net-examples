// Title: Aspose.Cells C# – Set Chart Category Axis Using a String Array
// Description: Learn how to assign custom X‑axis labels to an Aspose.Cells column chart by converting a C# string[] into the required CategoryData format and saving the workbook as XLSX.
// Keywords: Aspose.Cells chart category axis | C# set X axis labels | CategoryData string array | column chart custom labels Aspose | .NET chart category data | Aspose.Cells NSeries CategoryData | chart axis from array | Aspose.Cells example C# | Excel chart custom categories
// Common Searches: Aspose.Cells set chart X axis from string array | C# assign custom category labels to Excel chart | How to use NSeries.CategoryData with string[] | Aspose.Cells column chart custom categories example | Convert string array to CategoryData format
// Developer Intent: Apply a predefined string array as the category (X‑axis) labels of an Aspose.Cells chart.
// Use Cases: Create a column chart where the X‑axis shows product names, status codes, or any textual categories stored in memory. | Dynamically generate chart labels from user input, a database query, or an API response without writing them to worksheet cells. | Export reports that require non‑numeric axis labels, such as quarterly phases, department names, or custom tags.
// AI Prompts: Show C# code that sets an Aspose.Cells chart's category axis from a string[] array. | Generate a snippet converting a string array to the CategoryData format required by Aspose.Cells and applying it to a chart series. | Explain step‑by‑step how to assign custom X‑axis labels to an Aspose.Cells NSeries.CategoryData property in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Learn how to assign custom X‑axis labels to an Aspose.Cells column chart by converting a C# string[] into the required CategoryData format and saving the workbook as XLSX.
class SetCategoryAxisFromArray
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some numeric data for the chart series
        worksheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            worksheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the series data range (Y‑values)
        chart.NSeries.Add("B2:B5", true);

        // Define the category names in a string array
        string[] categories = new string[] { "Alpha", "Beta", "Gamma", "Delta" };

        // Convert the array to the format expected by Aspose.Cells:
        // {"Alpha","Beta","Gamma","Delta"}
        string categoryData = "{" + string.Join(",", categories) + "}";

        // Assign the category data to the chart's X‑axis
        chart.NSeries.CategoryData = categoryData;

        // Save the workbook
        workbook.Save("ChartWithStringArrayCategories.xlsx");
    }
}
