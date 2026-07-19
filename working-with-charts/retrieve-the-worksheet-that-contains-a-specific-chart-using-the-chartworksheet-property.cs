// Title: Get the Parent Worksheet of a Chart with Aspose.Cells Chart.Worksheet (C#)
// Description: Shows how to create a workbook, add a column chart, and retrieve the worksheet that hosts the chart using the Chart.Worksheet property in Aspose.Cells for .NET. The code prints the sheet name and index before saving the file.
// Keywords: Aspose.Cells | Chart.Worksheet | C# | .NET | retrieve chart worksheet | parent worksheet of chart | chart worksheet property | Aspose.Cells chart example | get chart's worksheet
// Common Searches: Aspose.Cells get chart worksheet | Chart.Worksheet property C# | how to find which sheet a chart belongs to Aspose.Cells | retrieve parent worksheet of chart Aspose.Cells .NET | chart.Worksheet example
// Developer Intent: Obtain the worksheet that a specific chart resides on by accessing the Chart.Worksheet property.
// Use Cases: After adding a chart, determine its parent sheet to rename, hide, or apply formatting. | Loop through all charts in a workbook and log each chart’s worksheet name and index. | Validate chart placement before exporting or sharing the workbook. | Programmatically move or copy a chart based on its originating worksheet.
// AI Prompts: Write C# code using Aspose.Cells to add a column chart and then retrieve its parent worksheet via Chart.Worksheet, printing the sheet name and index. | Generate an example that iterates over every chart in a workbook and outputs each chart’s worksheet name and index using the Chart.Worksheet property. | Explain the purpose of the Chart.Worksheet property in Aspose.Cells and when it is preferable to use it instead of searching the Worksheets collection.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add a column chart, and retrieve the worksheet that hosts the chart using the Chart.Worksheet property in Aspose.Cells for .NET. The code prints the sheet name and index before saving the file.
class RetrieveChartWorksheet
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Retrieve the worksheet that contains this chart using Chart.Worksheet property
        Worksheet chartWorksheet = chart.Worksheet;
        Console.WriteLine("Chart belongs to worksheet: " + chartWorksheet.Name + " (Index: " + chartWorksheet.Index + ")");

        // Save the workbook
        workbook.Save("RetrieveChartWorksheet_out.xlsx");
    }
}
