// Title: Configure the X‑axis as a CategoryScale for a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a workbook, adds a column chart, assigns text labels to the X‑axis, and sets the CategoryAxis.CategoryType to CategoryScale. | Modify an existing Aspose.Cells chart so that the X‑axis displays sequential text categories by changing its axis type to CategoryScale and adding an axis title.
// Common Searches: aspocells c# set chart x axis to category type for column chart | display text labels on chart x axis with Aspose.Cells .NET | change chart category axis to treat labels as categories in C# | C# Aspose.Cells column chart with categorical x axis | how to configure chart axis as category in Aspose.Cells
// Tags: Aspose.Cells X axis category type | C# column chart category axis labels | Aspose.Cells chart axis configuration | Excel chart text categories .NET | Category axis scaling Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells with category names and numeric values, adds a column chart, links the series and category data ranges, sets the X‑axis (CategoryAxis) to CategoryScale so the labels are treated as sequential text categories, assigns a title to the axis, and saves the workbook as an XLSX file.
class ChangeXAxisToCategory
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));   // Text labels for X axis
            sheet.Cells[$"B{i}"].PutValue(i * 10);            // Corresponding numeric values
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

        // Change the X axis (category axis) type to CategoryScale
        // This ensures the axis treats the labels as sequential text categories
        chart.CategoryAxis.CategoryType = CategoryType.CategoryScale;

        // Optional: give the axis a title
        chart.CategoryAxis.Title.Text = "Categories";

        // Save the workbook to a file
        workbook.Save("ChartWithCategoryXAxis.xlsx", SaveFormat.Xlsx);
    }
}
