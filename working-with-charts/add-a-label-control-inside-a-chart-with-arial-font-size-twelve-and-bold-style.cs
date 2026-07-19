// Title: Add an Inside Chart Label with Arial 12‑Bold Font Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate it with sample data, generate a column chart, insert a label inside the chart area with ShapeCollection.AddLabelInChart, and format the label text to Arial, 12‑point, bold using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart label | AddLabelInChart example | inside chart text Aspose.Cells | format chart label font | Arial 12 bold Aspose.Cells | Aspose.Cells chart annotation | C# Excel chart label
// Common Searches: how to add a label inside a chart with Aspose.Cells | set font of chart label Aspose.Cells C# | AddLabelInChart method usage | position chart label Aspose.Cells .NET | change chart label color Aspose.Cells
// Developer Intent: Insert a text label into the plot area of a chart and apply Arial 12‑point bold styling programmatically with Aspose.Cells for .NET.
// Use Cases: Add a descriptive label to a sales chart to highlight a target value. | Automate dashboard generation where each chart includes a bold annotation. | Create Excel reports that require custom in‑chart titles or notes without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells to place a chart label at (150,150) using Times New Roman, 14‑point, italic. | Explain step‑by‑step how to align a chart label to the top‑right corner and set its background color in Aspose.Cells. | Show how to retrieve an existing chart label and modify its font size and color in a workbook.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, populate it with sample data, generate a column chart, insert a label inside the chart area with ShapeCollection.AddLabelInChart, and format the label text to Arial, 12‑point, bold using Aspose.Cells for .NET.
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Add a label inside the chart using the ShapeCollection.AddLabelInChart method
        // Parameters are in units of 1/4000 of the chart area
        Label chartLabel = chart.Shapes.AddLabelInChart(
            top: 100,    // vertical offset
            left: 100,   // horizontal offset
            height: 200, // height of the label
            width: 200   // width of the label
        );

        // Set label text
        chartLabel.Text = "Sample Chart Label";

        // Configure the label's font: Arial, size 12, bold
        chartLabel.Font.Name = "Arial";
        chartLabel.Font.Size = 12;
        chartLabel.Font.IsBold = true;
        chartLabel.Font.Color = Color.Black;

        // Save the workbook
        workbook.Save("ChartWithLabel.xlsx");
    }
}
