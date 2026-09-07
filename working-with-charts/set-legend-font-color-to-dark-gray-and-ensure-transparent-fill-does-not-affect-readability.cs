// Title: Set chart legend entry font to dark gray and make legend background opaque using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and changes every legend entry’s font color to DarkGray while keeping the text fill enabled. | Update an existing Aspose.Cells chart so the legend text is non‑transparent and the legend’s background mode is set to Opaque for improved readability.
// Common Searches: how to change legend entry font color to dark gray in Aspose.Cells .NET | Aspose.Cells make chart legend background opaque | prevent transparent legend text in Excel chart using Aspose.Cells C# | set legend entry formatting dark gray font Aspose.Cells example
// Tags: legend entry font color Aspose.Cells | opaque legend background Aspose.Cells | chart legend text fill Aspose.Cells | disable legend transparency Aspose.Cells | column chart legend formatting .NET

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, enables the legend, then iterates through each LegendEntry to set the font color to DarkGray, ensures the text fill is not transparent, and makes the legend background opaque before saving the file as LegendFontDarkGray.xlsx.
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
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(90);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the legend is visible
        chart.ShowLegend = true;

        // Iterate through all legend entries and apply the required formatting
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            // Set the font color to dark gray
            entry.Font.Color = Color.DarkGray;

            // Ensure the text has a fill (not transparent) so it remains readable
            entry.IsTextNoFill = false;

            // Make the background opaque to avoid transparency issues
            entry.BackgroundMode = BackgroundMode.Opaque;
        }

        // Save the workbook
        workbook.Save("LegendFontDarkGray.xlsx");
    }
}
