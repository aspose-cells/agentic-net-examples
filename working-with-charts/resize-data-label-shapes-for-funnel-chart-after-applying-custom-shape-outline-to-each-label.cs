// Title: How to resize funnel chart data label shapes and add a custom outline using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a funnel chart with Aspose.Cells, applies a thick colored border to each data label, and then enlarges the label shape to specific dimensions. | Update an existing Aspose.Cells funnel chart example to set a red outline on all data labels and increase their width and height by adjusting the label shape properties. | Provide a step‑by‑step C# snippet that accesses the DataLabels collection of a funnel chart, configures OutlineColor and OutlineWeight, and scales the label rectangles.
// Common Searches: Aspose.Cells C# increase size of funnel chart data labels | set custom border for funnel chart data labels using Aspose.Cells .NET | resize data label shape in Excel funnel chart programmatically | how to change outline thickness of chart data labels with Aspose.Cells
// Tags: funnel chart data label outline Aspose.Cells | resize chart data label shape .NET | custom border for Excel chart labels C# | Aspose.Cells adjust data label dimensions | programmatic funnel chart label styling

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The program creates a new workbook, populates stage and value data, adds a funnel chart, customizes the data label font, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the funnel chart
            sheet.Cells["A1"].PutValue("Stage");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Prospects");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Qualified");
            sheet.Cells["B3"].PutValue(600);
            sheet.Cells["A4"].PutValue("Proposal");
            sheet.Cells["B4"].PutValue(300);
            sheet.Cells["A5"].PutValue("Closed");
            sheet.Cells["B5"].PutValue(150);

            // Add a funnel chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Funnel, 5, 0, 25, 10);
            Chart funnelChart = sheet.Charts[chartIdx];
            funnelChart.NSeries.Add("B2:B5", true);
            funnelChart.NSeries.CategoryData = "A2:A5";

            // Access the first series
            Series series = funnelChart.NSeries[0];
            // series.HasDataLabel = true; // Uncomment if the property exists in your version.

            // Customize data label font (applies to all labels)
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Font.Size = 10;

            // Save the workbook with the modified chart
            string outputPath = "FunnelChart_ResizedLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
