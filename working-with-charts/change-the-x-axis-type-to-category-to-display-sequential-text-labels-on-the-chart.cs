// Title: Aspose.Cells C# – Set X‑Axis to Category Scale for Text Labels in a Column Chart
// Description: Creates a new workbook, fills column A with sequential text labels and column B with numeric values, adds a column chart, binds the series and category ranges, switches the X‑axis to CategoryScale so labels are treated as categories, and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# chart axis | CategoryScale X axis | CategoryAxis CategoryType | column chart text labels | set chart X axis to category | .NET Excel chart example | Aspose.Cells CategoryType usage | Excel chart categorical axis | programmatic chart formatting
// Common Searches: Aspose.Cells set X axis to CategoryScale C# | how to display text labels on chart X axis Aspose | change chart axis type to category in .NET | C# Aspose.Cells column chart category axis example | programmatically set chart X axis as category
// Developer Intent: The developer needs to configure the chart’s X‑axis as a categorical scale so that sequential text labels appear correctly on the axis.
// Use Cases: Produce a sales column chart where product names are shown as X‑axis categories. | Generate a financial report chart with month names displayed on the X‑axis. | Export a bar chart that uses custom text categories (e.g., region codes) instead of numeric values.
// AI Prompts: Give me C# code to change a line chart’s X axis to CategoryScale using Aspose.Cells. | Show how to update category labels dynamically after setting the axis type to Category in a bar chart. | Explain the differences between CategoryScale, DateScale, and ValueScale axis types in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisDemo
{
    // Creates a new workbook, fills column A with sequential text labels and column B with numeric values, adds a column chart, binds the series and category ranges, switches the X‑axis to CategoryScale so labels are treated as categories, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1)); // sequential text labels
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and the category (X) axis
            chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

            // Change the X axis (category axis) type to CategoryScale
            chart.CategoryAxis.CategoryType = CategoryType.CategoryScale;

            // Save the workbook (lifecycle: save)
            workbook.Save("ChartWithCategoryAxis.xlsx", SaveFormat.Xlsx);
        }
    }
}
