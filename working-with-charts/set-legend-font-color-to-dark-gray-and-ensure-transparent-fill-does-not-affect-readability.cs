// Title: C# Aspose.Cells – Set Chart Legend Font to Dark Gray with Transparent Background
// Description: Creates a workbook, adds a column chart, shows the legend, then iterates through each LegendEntry to set the font color to DarkGray, disables text‑no‑fill, applies a transparent background, and saves the file as LegendFontDarkGray.xlsx.
// Keywords: Aspose.Cells | C# chart legend font color | dark gray legend | transparent legend background | LegendEntry styling | Aspose.Cells chart customization | set legend font color .NET | chart legend readability | BackgroundMode.Transparent | IsTextNoFill false
// Common Searches: Aspose.Cells change legend font color | set legend text color dark gray Aspose.Cells | transparent legend background Aspose.Cells C# | make chart legend readable Aspose.Cells | how to style legend entries in Aspose.Cells
// Developer Intent: Modify a chart legend so the font is dark gray while the background stays transparent and the text remains readable.
// Use Cases: Standardize report charts by applying a corporate dark‑gray font to all legend entries. | Keep legend text visible when using a transparent background in automated spreadsheet generation. | Create visually consistent legends for dashboards that require a clean, non‑filled background.
// AI Prompts: Generate C# code with Aspose.Cells that sets the legend font color to dark gray and ensures the text is not transparent. | Show how to loop through chart.Legend.LegendEntries to apply BackgroundMode.Transparent and disable IsTextNoFill. | Explain best practices for styling chart legends for readability when using transparent backgrounds in Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, shows the legend, then iterates through each LegendEntry to set the font color to DarkGray, disables text‑no‑fill, applies a transparent background, and saves the file as LegendFontDarkGray.xlsx.
class SetLegendFontColor
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
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(50);
        sheet.Cells["B4"].PutValue(40);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the legend is visible
        chart.ShowLegend = true;

        // Iterate through each legend entry and customize its appearance
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            // Set the font color to dark gray
            entry.Font.Color = Color.DarkGray;

            // Make sure the text has a fill (no transparent text)
            entry.IsTextNoFill = false;

            // Set the background of the legend entry to transparent
            entry.BackgroundMode = BackgroundMode.Transparent;
        }

        // Save the workbook
        workbook.Save("LegendFontDarkGray.xlsx");
    }
}
