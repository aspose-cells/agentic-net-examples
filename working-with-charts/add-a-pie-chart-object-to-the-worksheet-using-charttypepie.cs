// Title: Add a Pie chart to an Excel worksheet with Aspose.Cells for .NET (C# example)
// AI Prompts: Write C# code that creates a new workbook, populates category and value cells, inserts a Pie chart using ChartType.Pie, binds the series to B2:B4 and categories to A2:A4, and saves the file as an .xlsx with Aspose.Cells. | Demonstrate how to adjust the position and size of a Pie chart in an Aspose.Cells worksheet by specifying start and end rows and columns. | Provide a C# snippet that adds a title and legend to a Pie chart created with Aspose.Cells, then exports the workbook to a file named PieChart.xlsx.
// Common Searches: aspnet cells c# how to insert a pie chart into an existing worksheet | example of binding category labels to a pie chart using Aspose.Cells | aspose.cells create pie chart from cell range B2:B4 and A2:A4 | set pie chart dimensions rows 5 to 15 columns 0 to 5 Aspose.Cells | save workbook with pie chart as PieChart.xlsx using Aspose.Cells
// Tags: Aspose.Cells add pie chart C# | Aspose.Cells chart data binding | Aspose.Cells set chart position rows columns | Aspose.Cells export workbook to xlsx | Aspose.Cells customize pie chart title legend

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Creates a new workbook, fills A1:B4 with category and value data, adds a Pie chart covering rows 5‑15 and columns 0‑5, binds values from B2:B4 and categories from A2:A4, then saves the workbook as PieChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet (rows 5‑15, columns 0‑5)
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data series and category labels for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Save the workbook with the newly added chart
        workbook.Save("PieChart.xlsx", SaveFormat.Xlsx);
    }
}
