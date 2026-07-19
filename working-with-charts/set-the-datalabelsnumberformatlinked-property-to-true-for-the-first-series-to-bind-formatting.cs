// Title: Aspose.Cells .NET – Bind Data Labels Number Format to Linked Cells for the First Series in a Column Chart
// Description: Demonstrates how to create a workbook, add a column chart, link data labels to cells, and set DataLabels.NumberFormatLinked = true so the labels inherit the number format of the source cells, then save the file as XLSX.
// Keywords: Aspose.Cells DataLabels NumberFormatLinked | chart data label formatting .NET | link data labels to cells Aspose | column chart series formatting | inherit number format from source cells
// Common Searches: Aspose.Cells set DataLabels.NumberFormatLinked | bind chart data label format to linked cells .NET | how to link data label number format in Aspose.Cells | chart series data label formatting example | Aspose.Cells column chart data labels
// Developer Intent: Enable NumberFormatLinked for the first series so its data labels automatically use the number format of the linked source cells.
// Use Cases: Display values with custom units (e.g., "100 units") while keeping numeric formatting consistent. | Create charts where label formatting updates automatically when the source cell format changes. | Produce financial or KPI reports where data labels inherit currency or percentage formats from a separate column.
// AI Prompts: Show code to enable NumberFormatLinked for every series in an Aspose.Cells chart. | Explain how to toggle DataLabels.ShowValue and NumberFormatLinked based on user preferences in a .NET app. | Provide a step‑by‑step guide to conditionally apply NumberFormatLinked when linked cells contain different formats.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, link data labels to cells, and set DataLabels.NumberFormatLinked = true so the labels inherit the number format of the source cells, then save the file as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for categories, values and formatted values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);

        sheet.Cells["C1"].PutValue("Formatted");
        sheet.Cells["C2"].PutValue("100 units");
        sheet.Cells["C3"].PutValue("200 units");

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the series data range and category (X‑axis) data
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Access the first series (index 0)
        Series series = chart.NSeries[0];

        // Enable data labels and link them to the formatted cells
        series.DataLabels.ShowValue = true;
        series.DataLabels.LinkedSource = "C2:C3";

        // Bind the number format of the data labels to the linked cells
        series.DataLabels.NumberFormatLinked = true;

        // Save the workbook to an XLSX file
        workbook.Save("DataLabelsNumberFormatLinkedDemo.xlsx");
    }
}
