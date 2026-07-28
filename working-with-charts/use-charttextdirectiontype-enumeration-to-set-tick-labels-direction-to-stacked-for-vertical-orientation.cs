// Title: C# – Set Value Axis Tick Labels to Stacked Using ChartTextDirectionType in Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a column chart, and applies ChartTextDirectionType.Stacked to the ValueAxis.TickLabels.DirectionType so the vertical axis labels are displayed in a stacked orientation before saving the file.
// Keywords: Aspose.Cells | ChartTextDirectionType | stacked tick labels | value axis label orientation | .NET chart customization | C# Excel chart example
// Common Searches: Aspose.Cells set tick label direction stacked | ChartTextDirectionType usage C# | how to stack vertical axis labels in Aspose.Cells | C# change chart axis label orientation Aspose | stacked tick labels column chart Aspose.Cells
// Developer Intent: Apply the Stacked direction to the value‑axis tick labels of a chart using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports where long numeric labels on the value axis would overlap, improving readability by stacking them. | Standardize chart appearance across automated dashboards that require consistent label orientation. | Customize multi‑series column charts in financial or scientific workbooks to prevent label clipping on the vertical axis.
// AI Prompts: Show a C# example that sets both value and category axes to Stacked using ChartTextDirectionType in Aspose.Cells. | Explain the impact of each ChartTextDirectionType enum value on axis label rendering. | Provide a step‑by‑step guide to change tick label direction for a line chart in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, and applies ChartTextDirectionType.Stacked to the ValueAxis.TickLabels.DirectionType so the vertical axis labels are displayed in a stacked orientation before saving the file.
public class SetTickLabelsDirectionStacked
{
    public static void Run()
    {
        try
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set tick labels direction to Stacked for the vertical (value) axis
            chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Stacked;

            // Define output file path
            string outputPath = "TickLabelsStackedDirection.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Entry point for the application
public class Program
{
    public static void Main(string[] args)
    {
        SetTickLabelsDirectionStacked.Run();
    }
}
