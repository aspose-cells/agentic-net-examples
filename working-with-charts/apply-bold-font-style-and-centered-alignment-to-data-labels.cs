// Title: Bold and Center Chart Data Labels with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create an Excel workbook, add a column chart, display values in the first series, apply bold font styling, and center the text both horizontally and vertically on data labels using Aspose.Cells for .NET. The workbook is saved as BoldCenteredDataLabels.xlsx.
// Keywords: Aspose.Cells C# chart data label formatting | bold font chart labels Aspose.Cells | center alignment data labels .NET | Excel chart label styling Aspose | Aspose.Cells US developers
// Common Searches: how to make chart data labels bold in Aspose.Cells C# | center data label text in Excel chart using Aspose.Cells | apply bold and centered alignment to Aspose.Cells chart labels | set horizontal and vertical alignment for chart data labels .NET
// Developer Intent: Apply bold styling and center alignment to chart data labels in an Excel file using Aspose.Cells for .NET.
// Use Cases: Enhance readability of column‑chart values in financial dashboards. | Standardize label appearance across multiple charts in automated report generation. | Prepare presentation‑ready Excel files with uniformly styled data labels.
// AI Prompts: Generate C# code with Aspose.Cells that sets data label font to italic and right‑aligns the text for a line chart. | Show an example of applying custom font size and background color to pie‑chart data labels using Aspose.Cells for .NET. | Explain how to toggle visibility and change alignment of data labels for several series in a single chart with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create an Excel workbook, add a column chart, display values in the first series, apply bold font styling, and center the text both horizontally and vertically on data labels using Aspose.Cells for .NET. The workbook is saved as BoldCenteredDataLabels.xlsx.
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series' data labels
        DataLabels dataLabels = chart.NSeries[0].DataLabels;

        // Show the values in the data labels
        dataLabels.ShowValue = true;

        // Apply bold font style
        dataLabels.Font.IsBold = true;

        // Center the text horizontally (and optionally vertically)
        dataLabels.TextHorizontalAlignment = TextAlignmentType.Center;
        dataLabels.TextVerticalAlignment = TextAlignmentType.Center;

        // Apply the font settings to all child label nodes
        dataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("BoldCenteredDataLabels.xlsx");
    }
}
