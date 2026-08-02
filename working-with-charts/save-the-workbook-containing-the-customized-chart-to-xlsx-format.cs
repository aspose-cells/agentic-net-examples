// Title: Save a Workbook with a Customized Column Chart to XLSX using Aspose.Cells for .NET
// Description: Creates a new workbook, inserts sample data, adds a column chart, and saves the entire workbook—including the chart—to an XLSX file with Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# chart export | save workbook as XLSX | column chart Aspose.Cells | export chart to Excel | SaveFormat.Xlsx | programmatic Excel chart creation
// Common Searches: Aspose.Cells save chart to xlsx | C# export column chart to Excel with Aspose | How to save workbook containing chart using Aspose.Cells | Save Aspose.Cells workbook as XLSX file | Create and save chart in Excel via Aspose.Cells .NET
// Developer Intent: Export the workbook that contains a customized column chart to an XLSX file.
// Use Cases: Generate a sales report with a column chart and deliver it as an Excel file. | Automate a financial dashboard where charts are added programmatically and saved for downstream analysis. | Produce template workbooks with predefined charts, populate them with data, and output each instance as a separate XLSX document.
// AI Prompts: Write C# code that adds a line chart, sets a title, and saves the workbook to XLSX using Aspose.Cells. | Explain how to modify axis labels and legend properties of a chart before saving the workbook with Aspose.Cells. | Show an example of saving multiple worksheets, each containing a different chart type, into a single XLSX file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, inserts sample data, adds a column chart, and saves the entire workbook—including the chart—to an XLSX file with Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Save the workbook (including the customized chart) to XLSX format
        workbook.Save("CustomizedChart.xlsx", SaveFormat.Xlsx);
    }
}
