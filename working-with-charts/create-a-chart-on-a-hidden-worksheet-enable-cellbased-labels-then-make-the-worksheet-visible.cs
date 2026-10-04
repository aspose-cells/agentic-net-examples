// Title: Create a column chart on a hidden worksheet, link series to cell ranges, enable cell‑based labels, and unhide the sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a hidden worksheet, populates it with sample data, creates a column chart whose values and category labels are taken from cell ranges, then makes the worksheet visible and saves the workbook. | Show how to change the example to generate a line chart instead of a column chart while keeping the worksheet hidden until after the chart is created, using Aspose.Cells in C#.
// Common Searches: how to add a chart to a hidden worksheet with Aspose.Cells in C# | C# Aspose.Cells link chart series to specific cell ranges for values and categories | enable category labels from cells for a chart using Aspose.Cells .NET | unhide a worksheet after creating a chart programmatically with Aspose.Cells | save workbook after generating column chart on a hidden sheet Aspose.Cells
// Tags: Aspose.Cells hidden sheet chart generation | cell‑based labels for Aspose chart series | bind chart series to worksheet cells C# | make worksheet visible after chart Aspose | column chart from worksheet data Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates adding a hidden worksheet, filling it with data, creating a column chart whose series values and category labels are linked to cell ranges, setting a chart title, unhiding the worksheet, and saving the workbook as ChartOnHiddenSheet.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a hidden worksheet and obtain its reference
            Worksheet ws = workbook.Worksheets.Add("DataSheet");
            ws.IsVisible = false; // Hide the worksheet

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["A5"].PutValue("Apr");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(15);
            ws.Cells["B5"].PutValue(25);

            // Add a column chart to the hidden worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Link series to cell ranges for values and categories
            chart.NSeries.Add("DataSheet!B2:B5", true);               // Values
            chart.NSeries[0].XValues = "DataSheet!A2:A5";            // Category labels

            // Optional chart title
            chart.Title.Text = "Monthly Sales";

            // Make the worksheet visible after chart creation
            ws.IsVisible = true;

            // Save the workbook
            workbook.Save("ChartOnHiddenSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
