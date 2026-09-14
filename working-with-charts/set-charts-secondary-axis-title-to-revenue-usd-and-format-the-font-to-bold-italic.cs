// Title: Set a bold‑italic secondary Y‑axis title on a column chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart, adds a secondary value axis, sets its title to "Revenue (USD)", makes the title visible, and applies bold and italic styling using Aspose.Cells. | Modify an existing Aspose.Cells workbook to display a secondary Y‑axis title with bold‑italic font on a chart, ensuring the title text is "Revenue (USD)".
// Common Searches: Aspose.Cells C# how to add a secondary axis title to a chart | set secondary Y axis label bold italic in Aspose.Cells workbook | C# Aspose.Cells column chart with secondary value axis title example | format secondary axis title font Aspose.Cells .NET | make secondary axis title visible Aspose.Cells chart C#
// Tags: Aspose.Cells set secondary axis label | C# format axis title font bold italic | column chart secondary value axis Aspose.Cells | Excel chart axis customization .NET | add secondary Y axis title Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart with primary and secondary series, plots the second series on a secondary Y‑axis, sets the secondary axis title to "Revenue (USD)", makes it visible, applies bold‑italic formatting to the title font, and saves the workbook.
class SetSecondaryAxisTitle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Primary");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        sheet.Cells["C1"].PutValue("Secondary");
        sheet.Cells["C2"].PutValue(5000);
        sheet.Cells["C3"].PutValue(7000);
        sheet.Cells["C4"].PutValue(6500);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add the primary series
        chart.NSeries.Add("B2:B4", true);
        // Add the secondary series and plot it on the secondary axis
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Set category (X) axis data
        chart.NSeries.CategoryData = "A2:A4";

        // Access the secondary (second Y) axis
        Axis secondaryAxis = chart.SecondValueAxis;

        // Set the title text and make it visible
        secondaryAxis.Title.Text = "Revenue (USD)";
        secondaryAxis.Title.IsVisible = true;

        // Format the title font: bold and italic
        secondaryAxis.Title.Font.IsBold = true;
        secondaryAxis.Title.Font.IsItalic = true;

        // Save the workbook
        workbook.Save("ChartWithSecondaryAxisTitle.xlsx");
    }
}
