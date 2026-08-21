// Title: C# – Set X‑Axis to Category Scale for Text Labels in an Aspose.Cells Column Chart
// Description: Creates a workbook, fills column A with text categories and column B with numeric values, adds a column chart, binds the series and category ranges, changes the X‑axis to CategoryScale via Chart.CategoryAxis.CategoryType, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# chart axis category | CategoryScale | text labels on X axis | column chart example | Excel export | CategoryAxis.CategoryType
// Common Searches: Aspose.Cells set X axis to category scale C# | how to display text labels on chart X axis Aspose.Cells | C# change chart axis type to CategoryScale | Aspose.Cells column chart category axis settings | example of CategoryType.CategoryScale in Aspose.Cells
// Developer Intent: Configure the chart’s X‑axis as a category scale so that sequential text labels appear instead of numeric values.
// Use Cases: Generate a sales‑by‑region column chart where each region name is shown on the X‑axis. | Create a marketing dashboard Excel file with custom category labels aligned to data points. | Update an existing Aspose.Cells chart to use text categories without rebuilding the chart.
// AI Prompts: Write C# code using Aspose.Cells to set a chart’s X‑axis to CategoryScale and bind custom text labels. | Explain the impact of CategoryType.CategoryScale on axis rendering in Aspose.Cells charts. | Show how to modify the X‑axis type of an existing Aspose.Cells column chart to CategoryScale programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills column A with text categories and column B with numeric values, adds a column chart, binds the series and category ranges, changes the X‑axis to CategoryScale via Chart.CategoryAxis.CategoryType, and saves the result as an XLSX file.
class ChangeXAxisToCategory
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (categories as text and numeric values)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X‑axis) labels
        chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

        // Change the X axis (category axis) type to CategoryScale
        chart.CategoryAxis.CategoryType = CategoryType.CategoryScale;

        // Save the workbook
        workbook.Save("ChartWithCategoryXAxis.xlsx", SaveFormat.Xlsx);
    }
}
