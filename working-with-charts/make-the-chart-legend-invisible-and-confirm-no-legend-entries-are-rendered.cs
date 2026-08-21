// Title: Hide Chart Legend in Aspose.Cells for .NET (C#) and Verify Entries
// Description: Creates a workbook, adds a column chart with sample data, disables the legend using chart.ShowLegend = false, outputs the ShowLegend flag and the count of LegendEntries (which remain in the collection), and saves the result as ChartWithoutLegend.xlsx.
// Keywords: Aspose.Cells hide legend | chart.ShowLegend false | C# Aspose.Cells chart legend | verify legend entries Aspose.Cells | remove chart legend .NET
// Common Searches: Aspose.Cells hide chart legend C# | chart.ShowLegend property example | check legend entries after hiding legend Aspose.Cells | remove legend from Excel chart using Aspose.Cells | how to make chart legend invisible Aspose.Cells
// Developer Intent: Make the chart legend invisible and confirm that no legend entries are rendered in the output file.
// Use Cases: Produce a clean column chart for dashboards where a legend is unnecessary. | Create printable Excel charts without overlapping legend text. | Hide legends before exporting workbooks to PDF to reduce visual clutter.
// AI Prompts: Generate C# code with Aspose.Cells that hides a chart legend and logs the number of legend entries after disabling it. | Explain the effect of setting chart.ShowLegend = false on the Legend.LegendEntries collection and how to confirm the legend is not displayed in the saved workbook. | Suggest alternative methods to remove legend entries from an Aspose.Cells chart without using the ShowLegend property.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendVisibility
{
    // Creates a workbook, adds a column chart with sample data, disables the legend using chart.ShowLegend = false, outputs the ShowLegend flag and the count of LegendEntries (which remain in the collection), and saves the result as ChartWithoutLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend
            chart.ShowLegend = false;

            // Verify that the legend is hidden and that no legend entries will be rendered
            // (The entries collection still exists, but they will not be displayed because ShowLegend is false)
            Console.WriteLine("Chart.ShowLegend = " + chart.ShowLegend);
            Console.WriteLine("Number of legend entries (still present in collection): " + chart.Legend.LegendEntries.Count);

            // Save the workbook
            workbook.Save("ChartWithoutLegend.xlsx");
        }
    }
}
