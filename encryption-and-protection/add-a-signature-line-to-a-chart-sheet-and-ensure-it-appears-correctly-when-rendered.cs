// Title: Add a signature line shape to an Excel chart sheet with Aspose.Cells for .NET (C#)
// Common Searches: how to insert a line shape as a signature on a chart sheet using Aspose.Cells C# | Aspose.Cells C# add line shape to chart sheet and save workbook | rendering a signature line on an Excel chart sheet with Aspose.Cells .NET | create chart sheet with column chart and custom line shape Aspose.Cells | C# example for adding a signature line to an Excel chart sheet using Aspose.Cells
// Tags: add line shape to chart sheet Aspose.Cells | signature line shape Excel chart sheet C# | Aspose.Cells chart sheet shape rendering | column chart with custom line shape Aspose.Cells | set shape coordinates Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a new workbook, adds a worksheet used as a chart sheet, populates sample data, inserts a column chart, adds a line shape named "SignatureLine1" to serve as a signature line, and saves the file as ChartSheetWithSignature.xlsx.
class AddSignatureToChartSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet that will serve as a chart sheet
            int chartSheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "ChartSheet";

            // Populate sample data for the chart
            chartSheet.Cells["A1"].PutValue("Category");
            chartSheet.Cells["B1"].PutValue("Value");
            chartSheet.Cells["A2"].PutValue("Jan");
            chartSheet.Cells["A3"].PutValue("Feb");
            chartSheet.Cells["A4"].PutValue("Mar");
            chartSheet.Cells["B2"].PutValue(120);
            chartSheet.Cells["B3"].PutValue(150);
            chartSheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet (acts as a chart sheet)
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = chartSheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            // Category data can be set if supported; omitted for compatibility with older versions
            // chart.NSeries[0].CategoryData = "A2:A4";

            chart.Title.Text = "Sales by Month";

            // Add a line shape to represent a signature line
            Shape signatureLine = chartSheet.Shapes.AddLine(30, 0, 10, 10, 100, 300);
            signatureLine.Name = "SignatureLine1";

            // Save the workbook
            string outputPath = "ChartSheetWithSignature.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
