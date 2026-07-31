// Title: C# – Create a Workbook, Populate Data, and Insert a Column Chart with Aspose.Cells
// Description: Demonstrates how to instantiate a Workbook, write headers and sample rows, add a Column chart using ChartCollection.Add, define its data range with SetChartDataRange, set a title, and save the file as an XLSX document.
// Keywords: Aspose.Cells C# column chart example | add chart to worksheet Aspose.Cells | ChartCollection.Add Aspose.Cells | SetChartDataRange Aspose.Cells | save workbook as XLSX Aspose.Cells | populate Excel cells programmatically
// Common Searches: how to add a column chart in Aspose.Cells C# | Aspose.Cells set chart data range | save Aspose.Cells workbook with chart | C# example for creating charts with Aspose.Cells | Aspose.Cells chart positioning rows columns
// Developer Intent: Generate an Excel file that contains sample data and a column chart using the Aspose.Cells .NET API.
// Use Cases: Build a sales‑by‑category report that visualizes values with a column chart. | Automate monthly performance dashboards that include pre‑formatted charts. | Create a reusable Excel template with a placeholder column chart for downstream data entry.
// AI Prompts: Modify the example to use a stacked column chart and change the chart title. | Add data labels to each column in the chart created with Aspose.Cells. | Reposition the chart to a different cell range and adjust its width and height programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to instantiate a Workbook, write headers and sample rows, add a Column chart using ChartCollection.Add, define its data range with SetChartDataRange, set a title, and save the file as an XLSX document.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: Workbook constructor)
        Workbook workbook = new Workbook();

        // Access the first worksheet (lifecycle rule: Worksheets property)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 5; i++)
        {
            sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
            sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet (feature rule: ChartCollection.Add)
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (vertical series)
        chart.SetChartDataRange("A1:B6", true);

        // Set a title for the chart (optional customization)
        chart.Title.Text = "Sample Column Chart";

        // Save the workbook (lifecycle rule: Workbook.Save)
        workbook.Save("ColumnChartDemo.xlsx", SaveFormat.Xlsx);
    }
}
