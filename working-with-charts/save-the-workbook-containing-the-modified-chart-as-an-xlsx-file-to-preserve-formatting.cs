// Title: C# Example: Save Workbook with Modified Chart as XLSX Using Aspose.Cells
// Description: Shows how to create a workbook, fill it with sample data, add a column chart, set a title, and save the file as XLSX with Aspose.Cells for .NET so that all chart formatting and layout are preserved.
// Keywords: Aspose.Cells | C# | .NET | save chart as xlsx | preserve chart formatting | export workbook with chart | column chart Aspose.Cells | Aspose.Cells SaveFormat.Xlsx | Excel chart generation C#
// Common Searches: Aspose.Cells save workbook with chart | C# export chart to XLSX using Aspose.Cells | preserve Excel chart formatting when saving | how to save modified chart as .xlsx in .NET | Aspose.Cells example chart save format
// Developer Intent: Save a workbook that contains a newly added or edited chart in XLSX format to retain all visual styling.
// Use Cases: Generate a sales report with a column chart and distribute it as an XLSX file. | Automate creation of dynamic charts from database values and keep the layout intact when saving. | Start from a template workbook, programmatically modify charts, and export the final version for downstream processing.
// AI Prompts: Write C# code with Aspose.Cells to add a pie chart, set its title, and save as XLSX while keeping formatting. | Explain how to change axis labels and legend position of a chart in Aspose.Cells before saving the workbook. | Provide a script that loops through multiple worksheets, adds a line chart to each, and saves each workbook as a separate XLSX file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, fill it with sample data, add a column chart, set a title, and save the file as XLSX with Aspose.Cells for .NET so that all chart formatting and layout are preserved.
public class SaveChartWorkbook
{
    public static void Main()
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

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Optional: set a chart title
        chart.Title.Text = "Fruit Sales";

        // Save the workbook as XLSX to preserve all formatting, including the chart
        workbook.Save("ModifiedChart.xlsx", SaveFormat.Xlsx);
    }
}
