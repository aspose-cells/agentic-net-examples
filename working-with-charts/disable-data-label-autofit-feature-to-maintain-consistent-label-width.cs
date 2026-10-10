// Title: Disable automatic data label width fitting for a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, shows data labels, and turns off the automatic width fitting of those labels. | Generate a C# Aspose.Cells snippet that sets DataLabels.AutoFit to false for a chart series and saves the workbook.
// Common Searches: how to turn off data label auto fit in Aspose.Cells C# chart | Aspose.Cells keep data label width constant in Excel column chart | C# Aspose.Cells disable automatic resizing of chart data labels | set DataLabels.AutoFit false Aspose.Cells .NET example
// Tags: Aspose.Cells disable data label auto fit | C# chart data label fixed width | Aspose.Cells column chart label settings | DataLabels.AutoFit property .NET | Excel chart label width consistency Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates it with sample data, adds a column chart, enables data labels, disables the automatic width fitting of those labels, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart (from row 5, column 0 to row 20, column 10)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIdx];

            // Bind the values to the chart (B2:B4). The first argument is the data range,
            // the second argument indicates that the series are in columns.
            chart.NSeries.Add("B2:B4", true);

            // Show data labels (values) on the chart
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
