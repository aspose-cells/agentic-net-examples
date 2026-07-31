// Title: Create a Column Chart in Aspose.Cells for .NET by Populating Worksheet Data
// Description: Shows how to write headers and values to cells A1:C4, add a column chart, assign the data range (including headers), and save the workbook as ChartWithData.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells column chart C# | populate worksheet data Aspose.Cells | set chart data range .NET | add chart to Excel workbook | save Excel file with chart | Aspose.Cells chart automation | Excel chart source data programmatic
// Common Searches: Aspose.Cells add column chart from cells | C# set chart data range including headers Aspose.Cells | populate Excel worksheet for chart Aspose.Cells .NET | save workbook with chart using Aspose.Cells | how to create chart programmatically in Aspose.Cells
// Developer Intent: Programmatically fill worksheet cells and generate a column chart that uses that data.
// Use Cases: Create a sales dashboard where months are categories and product sales are plotted as columns. | Automate a performance report that writes KPI values to Excel and visualizes them with a chart on the same sheet. | Export database query results to an Excel file that is ready for immediate chart‑based analysis.
// AI Prompts: Generate C# code to add multiple series to a column chart after populating data in Aspose.Cells. | Provide examples of formatting the column chart (title, axis labels, colors) after setting the data range. | Explain how to calculate the chart data range dynamically based on the number of populated rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to write headers and values to cells A1:C4, add a column chart, assign the data range (including headers), and save the workbook as ChartWithData.xlsx using C# and Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with source data for the chart
        // Header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        // Data rows
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(20);

        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["C3"].PutValue(40);

        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(50);
        sheet.Cells["C4"].PutValue(60);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (including headers)
        chart.SetChartDataRange("A1:C4", true);

        // Save the workbook to a file
        workbook.Save("ChartWithData.xlsx");
    }
}
