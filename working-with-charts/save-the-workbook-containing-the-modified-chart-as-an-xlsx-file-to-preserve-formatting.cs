// Title: Save a Modified Chart to XLSX with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, generate a column chart, customize its title, and save the file as XLSX using Aspose.Cells so the chart formatting is retained.
// Keywords: Aspose.Cells save chart XLSX | preserve chart formatting .NET | export modified chart workbook | Aspose.Cells set chart title | C# Aspose.Cells chart example
// Common Searches: Aspose.Cells save workbook with chart as XLSX | preserve chart formatting when exporting to XLSX | C# Aspose.Cells modify chart title and save | how to keep chart styles in Aspose.Cells XLSX output | save modified chart workbook Aspose.Cells .NET
// Developer Intent: Programmatically save a workbook that contains a customized chart in XLSX format while keeping all visual formatting intact.
// Use Cases: Generate a sales report, add a column chart, change the title, and distribute the XLSX file with exact visual layout. | Update existing chart elements (title, series, categories) in an automated workflow and export the result without losing styling. | Create multiple workbooks with different chart types, apply formatting program‑matically, and batch‑save them as XLSX for downstream analytics.
// AI Prompts: Write C# code using Aspose.Cells to add a pie chart, set the legend position, and save the workbook as XLSX while preserving all chart styles. | Provide a step‑by‑step guide to change axis labels of a chart in an Aspose.Cells workbook and export it to XLSX without losing formatting. | Generate a sample program that creates a line chart, applies a custom theme, and saves the workbook as XLSX using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample data, generate a column chart, customize its title, and save the file as XLSX using Aspose.Cells so the chart formatting is retained.
class SaveModifiedChart
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Cherry");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["B4"].PutValue(25);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Example modification: set a custom title
        chart.Title.Text = "Fruit Sales";

        // Save the workbook as XLSX to preserve formatting
        workbook.Save("ModifiedChart.xlsx", SaveFormat.Xlsx);
    }
}
