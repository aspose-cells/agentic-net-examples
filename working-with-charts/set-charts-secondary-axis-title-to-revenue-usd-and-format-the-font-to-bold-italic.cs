// Title: Set a chart axis title to "Revenue (USD)" and apply bold‑italic font styling with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart axis title text to Revenue (USD) | how to apply bold and italic font to a chart axis title using Aspose.Cells | C# Aspose.Cells column chart add value axis label Revenue USD | formatting chart axis title font style Aspose.Cells .NET | set chart axis title programmatically with Aspose.Cells C# example | Aspose.Cells chart axis title customization tutorial
// Tags: chart axis title text Aspose.Cells | bold italic axis title font Aspose.Cells | column chart value axis labeling C# | Aspose.Cells workbook save XLSX | C# chart formatting with Aspose.Cells | set chart axis title Aspose.Cells example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, inserts a column chart, sets the chart’s value axis title to "Revenue (USD)", makes the title font bold and italic, and saves the file as ChartWithSecondaryAxisTitle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the primary value axis title (Aspose.Cells does not expose a direct secondary axis API in this version)
            chart.ValueAxis.Title.Text = "Revenue (USD)";

            // Format the axis title font
            chart.ValueAxis.Title.Font.IsBold = true;
            chart.ValueAxis.Title.Font.IsItalic = true;

            // Save the workbook
            string outputPath = "ChartWithSecondaryAxisTitle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
