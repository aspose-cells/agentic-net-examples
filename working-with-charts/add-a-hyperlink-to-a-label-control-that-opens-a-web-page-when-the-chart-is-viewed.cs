// Title: Add a Hyperlink to a Chart Label with Aspose.Cells for .NET (C#)
// Description: C# example that creates an Excel workbook using Aspose.Cells, adds a column chart, inserts a label inside the chart area, and attaches a clickable hyperlink that opens a web page when the chart is viewed.
// Keywords: Aspose.Cells | C# chart label hyperlink | Excel chart hyperlink Aspose | Add label to chart Aspose.Cells | Hyperlink on chart shape .NET | Aspose.Cells chart example | Excel hyperlink label code | Aspose.Cells GitHub example
// Common Searches: Aspose.Cells add hyperlink to chart label C# | How to make chart label clickable in Excel using Aspose.Cells | C# code for hyperlink on chart shape Aspose | Aspose.Cells chart label link example | Set URL on Excel chart label .NET
// Developer Intent: Create a chart label in an Excel workbook and bind a hyperlink that opens a specified URL when the chart is displayed.
// Use Cases: Insert a promotional link inside a sales chart that directs users to a product landing page. | Provide a documentation URL within a chart label for quick access to help topics. | Embed a support contact link in a chart label so end‑users can request assistance. | Add a reference to an external data source from a financial chart label.
// AI Prompts: Generate C# code with Aspose.Cells that adds a label to a chart and assigns a hyperlink to https://www.aspose.com with a screen tip. | Explain how to adjust position, size, and hyperlink properties of a chart label in Aspose.Cells for .NET. | Show how to create multiple labeled hyperlinks across different charts in the same workbook using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// C# example that creates an Excel workbook using Aspose.Cells, adds a column chart, inserts a label inside the chart area, and attaches a clickable hyperlink that opens a web page when the chart is viewed.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
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

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);

            // Add a label inside the chart area (coordinates are in 1/4000 of the chart area)
            Label label = chart.Shapes.AddLabelInChart(1000, 1000, 500, 2000);
            label.Text = "Visit Aspose";
            label.Font.Color = Color.Blue;
            label.Font.Size = 12;

            // Configure hyperlink for the label using the existing Hyperlink object
            Hyperlink hyperlink = label.Hyperlink;
            hyperlink.Address = "https://www.aspose.com";
            hyperlink.TextToDisplay = "Aspose Website";
            hyperlink.ScreenTip = "Click to open Aspose";

            // Save the workbook
            string outputPath = "ChartWithLabelHyperlink.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
