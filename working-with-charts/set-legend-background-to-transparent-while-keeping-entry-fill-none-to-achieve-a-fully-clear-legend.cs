// Title: Set chart legend background to transparent and clear legend entry fill using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend background mode to transparent | remove fill from legend entries in an Aspose.Cells chart programmatically | make legend background fully transparent while keeping only text in Aspose.Cells .NET | how to clear legend entry fill for a column chart with Aspose.Cells | transparent legend with no fill for legend text Aspose.Cells C# example
// Tags: Aspose.Cells chart legend background transparency | Aspose.Cells legend entry no fill | C# Aspose.Cells set legend transparent | Aspose.Cells column chart legend styling | Aspose.Cells transparent legend background .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace LegendTransparentDemo
{
    // The example creates a workbook, adds a column chart with sample data, sets the legend's BackgroundMode to Transparent, forces chart calculation, iterates through each series to set LegendEntry.IsTextNoFill = true (removing fill from legend entries), and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            const string outputPath = "LegendTransparentDemo.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate sample data for the chart
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["B4"].PutValue(40);

            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(60);
            sheet.Cells["C4"].PutValue(35);

            // -------------------------------------------------
            // Add a column chart
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Make the legend background fully transparent
            // -------------------------------------------------
            chart.Legend.BackgroundMode = BackgroundMode.Transparent;

            // -------------------------------------------------
            // Ensure each legend entry has no fill (text only)
            // -------------------------------------------------
            // The legend entries are created automatically after the chart data is set.
            // Access the collection via chart.Legend.LegendEntries.
            // If the collection is empty (e.g., before calculation), force a calculation first.
            chart.Calculate();

            // Iterate over each series' legend entry and set IsTextNoFill = true
            foreach (Series series in chart.NSeries)
            {
                LegendEntry entry = series.LegendEntry;
                if (entry != null)
                {
                    entry.IsTextNoFill = true; // Text will have no fill
                }
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
