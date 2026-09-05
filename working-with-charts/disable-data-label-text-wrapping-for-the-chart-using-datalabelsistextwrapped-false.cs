// Title: How to disable text wrapping for data labels in an Aspose.Cells column chart using C#
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells, shows data label values, and sets DataLabels.IsTextWrapped to false. | Demonstrate how to access a chart series' DataLabels object and turn off text wrapping in Aspose.Cells for .NET. | Provide a complete example that saves the workbook after disabling data label text wrapping in an Excel chart.
// Common Searches: Aspose.Cells C# chart data label wrap off | Set DataLabels.IsTextWrapped false in Aspose.Cells column chart | Disable text wrapping for Excel chart data labels using Aspose.Cells .NET | How to prevent data label text from wrapping in Aspose.Cells charts | C# Aspose.Cells example for chart data label formatting
// Tags: Aspose.Cells chart data label no wrap | C# DataLabels.IsTextWrapped usage | column chart data label formatting Aspose.Cells | Excel chart data label text wrap control | Aspose.Cells disable data label wrapping

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, enables data labels, disables their text wrapping with DataLabels.IsTextWrapped = false, and saves the file as ChartDataLabels_NoWrap.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the data labels of the first series and enable them
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Disable text wrapping for the data labels
        dataLabels.IsTextWrapped = false;

        // Save the workbook to a file
        workbook.Save("ChartDataLabels_NoWrap.xlsx");
    }
}
