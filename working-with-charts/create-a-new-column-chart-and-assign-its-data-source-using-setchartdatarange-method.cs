// Title: Aspose.Cells for .NET – Create a Column Chart and Set Its Data Source with SetChartDataRange (C#)
// Description: C# code that creates a workbook, fills cells A1:B4 with category and value data, adds a Column chart, binds the chart to that range using Chart.SetChartDataRange (plot by column), sets a chart title, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells C# SetChartDataRange | column chart Aspose.Cells | bind chart to range .NET | Chart.SetChartDataRange example | Aspose.Cells chart API | save workbook XLSX Aspose | GitHub Aspose.Cells chart sample | global
// Common Searches: Aspose.Cells SetChartDataRange C# example | how to bind data to a column chart in Aspose.Cells | create column chart programmatically with Aspose.Cells .NET | set chart title Aspose.Cells C# | save Aspose.Cells workbook as XLSX
// Developer Intent: Add a column chart, link it to a specific cell range, customize the title, and export the workbook programmatically.
// Use Cases: Generate a sales‑by‑region column chart directly from worksheet data for a monthly report. | Allow users to select a dynamic range in a UI and refresh the chart automatically using SetChartDataRange. | Create a pre‑formatted workbook with an embedded chart for automated distribution to stakeholders.
// AI Prompts: Write C# code that creates a stacked column chart with Aspose.Cells, sets its data range to A1:C5, and adds a legend. | Show how to change an existing column chart to a line chart after binding data with SetChartDataRange in Aspose.Cells. | Provide a snippet that reads a DataTable, writes it to a worksheet, and applies the values to a column chart using SetChartDataRange.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // C# code that creates a workbook, fills cells A1:B4 with category and value data, adds a Column chart, binds the chart to that range using Chart.SetChartDataRange (plot by column), sets a chart title, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Assign the data source to the chart using SetChartDataRange
            // The range includes both category (A1:A4) and values (B1:B4)
            // The second parameter 'true' indicates plotting by column
            chart.SetChartDataRange("A1:B4", true);

            // Optional: set a title for clarity
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to an XLSX file
            workbook.Save("ColumnChartWithDataRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
