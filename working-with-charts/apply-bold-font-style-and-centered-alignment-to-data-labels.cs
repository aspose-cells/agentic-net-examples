// Title: How to make chart data labels bold and center-aligned using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets chart series data labels to a bold font and centers the text horizontally. | Show how to apply font styling and horizontal alignment to column chart data labels in a .NET workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# set data label font to bold for column chart | center align data labels in Aspose.Cells chart programmatically | apply font style to chart series data labels using Aspose.Cells .NET | C# Aspose.Cells make data labels bold and centered in Excel file | formatting chart data labels with Aspose.Cells API
// Tags: bold font chart data labels Aspose.Cells | center horizontal alignment data labels Aspose.Cells | apply font styling to chart series Aspose.Cells C# | format chart data labels Excel Aspose.Cells | column chart label formatting Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, enables data labels, sets the labels to bold font and center alignment, applies the font settings, and saves the file as BoldCenteredDataLabels.xlsx.
class ApplyBoldCenteredDataLabels
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series and enable data labels
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Apply bold font style to the data labels
        series.DataLabels.Font.IsBold = true;

        // Center align the data label text horizontally
        series.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;

        // Propagate the font settings to all child label nodes
        series.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("BoldCenteredDataLabels.xlsx");
    }
}
