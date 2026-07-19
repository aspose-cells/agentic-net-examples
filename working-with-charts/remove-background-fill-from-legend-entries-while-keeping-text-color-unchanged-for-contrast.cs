// Title: Aspose.Cells C# – Remove Legend Background Fill While Preserving Text Color
// Description: Creates a workbook, adds a column chart, enables the legend, and sets each LegendEntry's BackgroundMode to Transparent so the legend background disappears but the font color stays unchanged. The file is saved as ChartWithoutLegendBackgroundFill.xlsx.
// Keywords: Aspose.Cells legend transparent background | C# chart legend entry fill removal | remove legend background Aspose.Cells | legend entry BackgroundMode Transparent | keep legend text color Aspose.Cells
// Common Searches: how to make chart legend transparent in Aspose.Cells C# | remove legend fill without changing font color Aspose.Cells | Aspose.Cells set legend entry background to none | transparent legend entries .NET Excel library
// Developer Intent: Eliminate the background fill of chart legend entries while leaving their text color intact for optimal contrast.
// Use Cases: Design Excel charts that overlay colored slide backgrounds without the legend obscuring the view. | Generate reports where cell shading is used and the legend must remain unobtrusive yet readable. | Standardize the appearance of legends across multiple chart types (column, pie, line) by applying a transparent background.
// AI Prompts: Write C# code with Aspose.Cells that loops through LegendEntryCollection and sets BackgroundMode to Transparent while preserving each entry's font color. | Show how to apply a transparent background to legend entries for different chart types (pie, line, bar) using Aspose.Cells in .NET. | Explain how to adjust legend text color for contrast after removing the background fill in an Aspose.Cells chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, enables the legend, and sets each LegendEntry's BackgroundMode to Transparent so the legend background disappears but the font color stays unchanged. The file is saved as ChartWithoutLegendBackgroundFill.xlsx.
class RemoveLegendBackgroundFill
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(70);
        sheet.Cells["C3"].PutValue(60);
        sheet.Cells["C4"].PutValue(90);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the legend is visible
        chart.ShowLegend = true;

        // Iterate through each legend entry and remove its background fill
        LegendEntryCollection entries = chart.Legend.LegendEntries;
        foreach (LegendEntry entry in entries)
        {
            // Set background mode to Transparent to remove fill
            entry.BackgroundMode = BackgroundMode.Transparent;

            // Keep the existing text color (no change needed)
            // If you want to enforce a specific contrast color, uncomment the line below:
            // entry.Font.Color = Color.Black;
        }

        // Save the workbook
        workbook.Save("ChartWithoutLegendBackgroundFill.xlsx");
    }
}
