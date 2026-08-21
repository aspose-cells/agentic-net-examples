// Title: Add a Bold Arial 12‑pt Label Inside a Chart with Aspose.Cells for .NET
// Description: Creates a workbook, fills sample data, inserts a column chart, and uses AddLabelInChart to place a label inside the chart area. The label text is set to "Chart Label" and formatted with Arial, 12‑pt, bold (optional black color), then saves the file as ChartWithLabel.xlsx.
// Keywords: Aspose.Cells | C# | AddLabelInChart | chart label | Arial 12 bold | text box in chart | column chart annotation | .NET Excel automation | US
// Common Searches: Aspose.Cells add label inside chart C# | Set chart label font Arial Aspose.Cells | Add text box to Excel chart .NET | How to format chart annotation Aspose.Cells | AddLabelInChart example
// Developer Intent: Place a formatted text label inside an Excel chart using Aspose.Cells for .NET.
// Use Cases: Add an annotation to highlight a data point within a column chart. | Insert a subtitle or note directly in the chart area for clearer reporting. | Apply corporate branding by styling chart labels with a specific font and size.
// AI Prompts: Generate C# code with Aspose.Cells that adds a label inside a chart and formats it as Arial 12‑pt bold. | Show how to position a label in a chart using AddLabelInChart with top, left, height, and width parameters. | Provide an example of adding multiple chart labels, each with different font styles, using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, fills sample data, inserts a column chart, and uses AddLabelInChart to place a label inside the chart area. The label text is set to "Chart Label" and formatted with Arial, 12‑pt, bold (optional black color), then saves the file as ChartWithLabel.xlsx.
class AddLabelToChart
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Add a label inside the chart (units are 1/4000 of the chart area)
        // top, left, height, width
        Label chartLabel = chart.Shapes.AddLabelInChart(100, 100, 200, 200);
        chartLabel.Text = "Chart Label";

        // Set label font: Arial, size 12, bold
        chartLabel.Font.Name = "Arial";
        chartLabel.Font.Size = 12;
        chartLabel.Font.IsBold = true;
        chartLabel.Font.Color = Color.Black; // optional color

        // Save the workbook
        workbook.Save("ChartWithLabel.xlsx");
    }
}
