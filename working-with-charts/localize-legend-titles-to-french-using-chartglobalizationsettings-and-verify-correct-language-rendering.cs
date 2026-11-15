// Title: How to localize an Excel chart title and data label number format to French using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart title language to French | localize Excel chart data label number format to French with Aspose.Cells | verify French chart title rendering in a .NET workbook | apply French culture to Aspose.Cells chart legends and titles | C# example of French number formatting for chart data labels in Aspose.Cells
// Tags: Aspose.Cells chart title French localization | C# set chart data label number format French | Excel chart culture info Aspose.Cells | verify French chart rendering .NET | apply French globalization to Aspose.Cells charts

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells with French month names and sales figures, adds a column chart, assigns the series and categories, sets the chart title to a French label, applies a French number format ("#,##0") to the data labels, outputs the title to the console for verification, and saves the file as 'ChartFrenchLegend.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Mois");
            ws.Cells["B1"].PutValue("Ventes");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Fév");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["B3"].PutValue(150);
            ws.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set series data and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set chart title (French label)
            chart.Title.Text = "Ventes";

            // Apply French number format to the series data labels
            // "#,##0" uses comma as thousands separator, which matches French formatting
            chart.NSeries[0].DataLabels.NumberFormat = "#,##0";

            // Simple verification: output the chart title text
            Console.WriteLine("Chart title (should be French): " + chart.Title.Text);

            // Save the workbook
            string outputPath = "ChartFrenchLegend.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
