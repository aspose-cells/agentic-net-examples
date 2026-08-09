// Title: Add a Pie Chart to an Excel Worksheet with Aspose.Cells (C#)
// Description: Demonstrates how to create a new Workbook, fill cells with category names and values, insert a Pie chart using ChartType.Pie, bind the series to B2:B4, set category labels from A2:A4, add a title, and save the file as PieChart.xlsx.
// Keywords: Aspose.Cells pie chart C# | ChartType.Pie example | add pie chart Aspose.Cells | C# Excel chart creation | Aspose.Cells save workbook with chart
// Common Searches: Aspose.Cells add pie chart .NET | C# create pie chart in Excel with Aspose | how to bind data to pie chart Aspose.Cells | set pie chart title Aspose.Cells C# | export Excel file with chart using Aspose
// Developer Intent: Insert a Pie chart, link it to worksheet data, customize the title, and export the workbook.
// Use Cases: Generate a product‑share pie chart for monthly sales reports. | Build an interactive dashboard that visualizes market segment distribution. | Export survey results with a summary pie chart of respondent choices.
// AI Prompts: Show how to display data labels on the pie chart in Aspose.Cells C#. | Provide code to change slice colors and explode a specific slice. | Explain how to adjust the chart range automatically when rows are added.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a new Workbook, fill cells with category names and values, insert a Pie chart using ChartType.Pie, bind the series to B2:B4, set category labels from A2:A4, add a title, and save the file as PieChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
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

        // Add a pie chart to the worksheet (topRow, leftColumn, bottomRow, rightColumn)
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Optional: set a title for the chart
        chart.Title.Text = "Fruit Distribution";

        // Save the workbook with the pie chart
        workbook.Save("PieChart.xlsx", SaveFormat.Xlsx);
    }
}
