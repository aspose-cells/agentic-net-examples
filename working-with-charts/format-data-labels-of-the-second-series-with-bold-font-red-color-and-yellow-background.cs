// Title: C# Aspose.Cells – Bold Red Font & Yellow Background for Second Series Data Labels in a Column Chart
// Description: Creates a workbook, adds two data series, inserts a column chart, enables data labels for the second series, then applies a bold typeface, red text color, and yellow label background before saving the file.
// Keywords: Aspose.Cells C# chart formatting | second series data label style | bold font Aspose.Cells | red text color chart label | yellow background data label | column chart Aspose.Cells | .NET Excel chart customization
// Common Searches: Aspose.Cells set data label font to bold C# | change chart label color to red Aspose.Cells | apply background color to data labels Aspose.Cells .NET | format second series labels in Excel chart using Aspose | C# Aspose.Cells chart label styling tutorial
// Developer Intent: Style the second series' data labels with bold red text on a yellow background.
// Use Cases: Emphasize a specific product line in a sales column chart | Apply corporate branding colors to a financial report's key metric | Distinguish a priority KPI in a performance dashboard | Create a presentation slide where one series stands out visually | Generate a printable Excel chart with highlighted data points
// AI Prompts: Write C# code using Aspose.Cells to make the second series data labels bold, red, and with a yellow background. | Explain step‑by‑step how to customize font style, text color, and label background for a particular series in an Aspose.Cells chart. | Show how to loop through chart series in Aspose.Cells and assign unique label formatting to each series.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds two data series, inserts a column chart, enables data labels for the second series, then applies a bold typeface, red text color, and yellow label background before saving the file.
class FormatSecondSeriesDataLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add both series to the chart
        chart.NSeries.Add("B2:B4", true); // First series
        chart.NSeries.Add("C2:C4", true); // Second series
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the second series
        Series secondSeries = chart.NSeries[1];
        secondSeries.DataLabels.ShowValue = true;

        // Apply bold font and red color to the data labels
        secondSeries.DataLabels.Font.IsBold = true;
        secondSeries.DataLabels.Font.Color = Color.Red;

        // Set yellow background for the data labels
        secondSeries.DataLabels.Area.BackgroundColor = Color.Yellow;

        // Apply the font settings to all child label nodes
        secondSeries.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("FormattedSecondSeriesDataLabels.xlsx");
    }
}
