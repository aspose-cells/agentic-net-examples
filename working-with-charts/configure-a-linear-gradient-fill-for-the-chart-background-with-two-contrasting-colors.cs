// Title: Create a column chart with a horizontal black‑to‑white two‑color linear gradient background using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel file that contains a column chart whose plot area background is a horizontal linear gradient from black to white using Aspose.Cells C# API. | Show how to apply a two‑color gradient fill to a chart's plot area by setting FillFormat.FillType to Gradient and calling SetTwoColorGradient with GradientStyleType.Horizontal in C#. | Write C# code that creates sample data, adds a column chart, and configures the chart background with a black‑to‑white linear gradient, then saves the workbook.
// Common Searches: Aspose.Cells C# set chart plot area background gradient | How to use SetTwoColorGradient for Excel chart background in .NET | Create column chart with black to white gradient using Aspose.Cells | Configure gradient variant for chart background Aspose.Cells .NET | Apply horizontal gradient style to chart background in C#
// Tags: Aspose.Cells chart plot area gradient fill | C# FillFormat SetTwoColorGradient usage | linear gradient fill for chart background | column chart background gradient Aspose.Cells | Excel workbook gradient fill API .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for FillFormat, FillType, GradientStyleType

// The example creates a workbook, adds sample data, inserts a column chart, accesses the chart's plot area FillFormat, sets a horizontal black‑to‑white two‑color linear gradient, and saves the workbook as an .xlsx file.
public class ChartBackgroundLinearGradient
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Item 1");
        sheet.Cells["A3"].PutValue("Item 2");
        sheet.Cells["A4"].PutValue("Item 3");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the plot area fill format (chart background)
        FillFormat backgroundFill = chart.PlotArea.Area.FillFormat;

        // Set gradient fill type
        backgroundFill.FillType = FillType.Gradient;

        // Apply a linear two‑color gradient (black to white, horizontal)
        backgroundFill.SetTwoColorGradient(
            Color.Black,               // Starting color
            Color.White,               // Ending color
            GradientStyleType.Horizontal,
            1);                        // Variant (1‑4)

        // Save the workbook
        string outputPath = "ChartBackgroundLinearGradient.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
