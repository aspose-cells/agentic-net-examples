// Title: Create a Workbook and Add a Column Chart with Aspose.Cells for .NET (C#)
// Description: C# example that creates a new Workbook, fills cells A1:B6 with sample data, adds a Column chart (rows 7‑20, columns 1‑8), sets the data range, gives the chart a title, and saves the file as ColumnChartDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# column chart example | add chart to worksheet Aspose.Cells | ChartCollection.Add ChartType.Column | SetChartDataRange Aspose.Cells | save workbook as xlsx Aspose | Aspose.Cells chart title | Excel automation .NET
// Common Searches: how to add a column chart in Aspose.Cells C# | Aspose.Cells set chart data range example | save Excel file with chart using Aspose.Cells | Aspose.Cells ChartCollection.Add usage | C# create Excel workbook with chart
// Developer Intent: Generate an Excel workbook, populate it with sample data, insert a column chart, and save the result programmatically with Aspose.Cells for .NET.
// Use Cases: Build a sales‑by‑category report that visualizes values in a column chart. | Automate monthly performance dashboards that include dynamically generated charts. | Export analytical data to XLSX files with ready‑to‑present column charts for stakeholders.
// AI Prompts: Show how to change the inserted chart to a stacked column chart in the same workbook. | Provide code to bind the chart to a named range instead of a hard‑coded cell range. | Explain how to add data labels and customize their font for the column chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // C# example that creates a new Workbook, fills cells A1:B6 with sample data, adds a Column chart (rows 7‑20, columns 1‑8), sets the data range, gives the chart a title, and saves the file as ColumnChartDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet (default worksheet is already present)
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 1; i <= 5; i++)
            {
                sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
                sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
            }

            // Add a column chart to the worksheet using ChartCollection.Add(ChartType, int, int, int, int)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 1, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart (vertical series)
            chart.SetChartDataRange("A1:B6", true);

            // Optional: set a title for the chart
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook (uses the Workbook.Save method rule)
            workbook.Save("ColumnChartDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
