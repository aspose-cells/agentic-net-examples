// Title: Aspose.Cells for .NET – C# – Set secondary axis title and apply bold italic font
// Description: This C# example creates a workbook, adds a column chart with primary and secondary series, plots the second series on the secondary value axis, then sets the secondary axis title to "Revenue (USD)" and formats the title text in bold and italic before saving the file.
// Keywords: Aspose.Cells | .NET | C# | chart | secondary axis | axis title | bold italic | font formatting | column chart | financial report | GitHub example
// Common Searches: Aspose.Cells set secondary axis title C# | Bold italic secondary axis title Aspose.Cells | How to format chart axis font in Aspose.Cells .NET | Add secondary value axis label Aspose.Cells chart | Aspose.Cells chart example GitHub
// Developer Intent: Add and style a secondary axis title in an Aspose.Cells chart using C#.
// Use Cases: Display revenue on a secondary axis with a clear, styled label in financial dashboards. | Differentiate primary and secondary data series in a column chart by using distinct axis titles. | Generate Excel reports programmatically where axis titles need bold‑italic emphasis for readability.
// AI Prompts: Write C# code with Aspose.Cells to set a secondary axis title and make the font bold and italic. | Show how to add a secondary value axis label and customize its font in an Aspose.Cells chart. | Explain steps to hide or show the secondary axis title and change its style in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, adds a column chart with primary and secondary series, plots the second series on the secondary value axis, then sets the secondary axis title to "Revenue (USD)" and formats the title text in bold and italic before saving the file.
class SetSecondaryAxisTitle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(150);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(250);

        sheet.Cells["C1"].PutValue("Revenue");
        sheet.Cells["C2"].PutValue(3000);
        sheet.Cells["C3"].PutValue(4000);
        sheet.Cells["C4"].PutValue(5000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first on primary axis, second on secondary axis
        chart.NSeries.Add("B2:B4", true);          // Sales series
        chart.NSeries.Add("C2:C4", true);          // Revenue series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Set the secondary axis title and format it
        chart.SecondValueAxis.Title.Text = "Revenue (USD)";
        chart.SecondValueAxis.Title.IsVisible = true;
        chart.SecondValueAxis.Title.Font.IsBold = true;
        chart.SecondValueAxis.Title.Font.IsItalic = true;

        // Save the workbook
        workbook.Save("SecondaryAxisTitle.xlsx");
    }
}
