// Title: Aspose.Cells for .NET – Set Chart Category Axis from a String Array (C#)
// Description: Learn how to assign custom X‑axis labels to an Aspose.Cells chart by building the required CategoryData string from a C# string array. The example creates a workbook, adds numeric series data, defines categories (Apple, Banana, Cherry), configures a column chart, sets NSeries.CategoryData, and saves the file.
// Keywords: Aspose.Cells chart category axis | C# CategoryData string array | custom X axis labels Aspose.Cells | Chart.NSeries.CategoryData example | Aspose.Cells for .NET chart labels | Excel chart custom categories C# | Aspose.Cells column chart string categories
// Common Searches: Aspose.Cells set chart X axis labels from array | C# Aspose.Cells CategoryData string format | How to use NSeries.CategoryData with string array | Aspose.Cells custom category axis example | Create chart with string categories in Aspose.Cells
// Developer Intent: Apply a string array as the category (X‑axis) labels of an Aspose.Cells chart.
// Use Cases: Display product names on the X‑axis of a sales column chart without writing them to worksheet cells. | Generate monthly reports where month names are supplied from a predefined list. | Standardize category labels across multiple charts in a single workbook using a reusable array.
// AI Prompts: Show a C# code snippet that sets Chart.NSeries.CategoryData from a string[] in Aspose.Cells. | Explain how to format the CategoryData sequence string for a chart when using an array of labels. | Provide an Aspose.Cells example that reads categories from a List<string> and applies them to a chart's X‑axis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryAxisExample
{
    // Learn how to assign custom X‑axis labels to an Aspose.Cells chart by building the required CategoryData string from a C# string array. The example creates a workbook, adds numeric series data, defines categories (Apple, Banana, Cherry), configures a column chart, sets NSeries.CategoryData, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample numeric data for the series
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Define the categories you want on the X‑axis
            string[] categories = new string[] { "Apple", "Banana", "Cherry" };
            // Build the sequence string that Aspose.Cells expects, e.g. {"Apple","Banana","Cherry"}
            string categoryData = "{" + string.Join(",", categories) + "}";

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data (values) and assign the custom category sequence
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = categoryData;   // <-- sets the category axis using the string array

            // Optional: give the series a name
            chart.NSeries[0].Name = "Sample Series";

            // Save the workbook
            workbook.Save("CategoryAxisWithStringArray.xlsx");
        }
    }
}
