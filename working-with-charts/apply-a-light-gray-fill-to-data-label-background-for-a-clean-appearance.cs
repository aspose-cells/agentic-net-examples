// Title: How to set a light gray solid fill for chart data label background using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells, enables data labels, and applies a solid light gray fill to the label background. | Write a snippet that uses Aspose.Cells Chart.NSeries.DataLabels to set the Area.FillFormat to solid and BackgroundColor to LightGray. | Provide a complete example that adds sample data, inserts a column chart, shows values on data labels, and customizes the label area color in an Excel workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# set data label background color to light gray | apply solid fill to chart data labels in .NET Excel file | change appearance of data labels in Aspose.Cells chart | C# Aspose.Cells chart data label area fill pattern example | customize data label background color in Excel using Aspose.Cells API
// Tags: Aspose.Cells chart data label background fill | C# chart label area fill style | Excel chart label appearance customization .NET | Aspose.Cells series label formatting | light gray data label area Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, enables data labels for the first series, and applies a solid light gray fill to the data label background before saving the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Apply a light gray solid fill to the data label background
        dataLabels.Area.FillFormat.Pattern = FillPattern.Solid;
        dataLabels.Area.BackgroundColor = Color.LightGray;

        // Save the workbook to a file
        workbook.Save("DataLabelsLightGrayBackground.xlsx");
    }
}
