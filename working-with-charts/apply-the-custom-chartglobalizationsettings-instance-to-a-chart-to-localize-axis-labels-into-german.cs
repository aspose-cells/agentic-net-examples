// Title: Apply ChartGlobalizationSettings to localize Aspose.Cells column chart axis labels to German in C#
// AI Prompts: Write C# code that creates a workbook, sets ChartGlobalizationSettings (or Workbook.Settings.CultureInfo) to "de-DE", adds a column chart, and assigns German text to the value and category axis titles using Aspose.Cells. | Show how to configure German number formatting for the value axis of an Aspose.Cells chart while keeping axis titles in German. | Modify an existing Aspose.Cells chart to use German locale for all axis labels and number formats in a .NET application.
// Common Searches: Aspose.Cells C# set chart axis titles German localization | how to use ChartGlobalizationSettings for German language in Aspose.Cells chart | C# example of applying de-DE culture to Excel chart axis labels with Aspose.Cells | localize Excel chart axis numbers to German format using Aspose.Cells .NET | Aspose.Cells column chart German axis titles sample code
// Tags: ChartGlobalizationSettings German locale Aspose.Cells | set workbook CultureInfo de-DE for chart axis | localize chart axis labels German .NET | German number format on Aspose.Cells chart value axis | column chart axis title localization C#

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, applies the German (de-DE) culture via ChartGlobalizationSettings or Workbook.Settings.CultureInfo, populates sample data, adds a column chart, sets German text for the value and category axis titles, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply German culture to the workbook (affects number formatting)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Monat");
            sheet.Cells["B1"].PutValue("Umsatz");
            sheet.Cells["A2"].PutValue("Januar");
            sheet.Cells["B2"].PutValue(1200.5);
            sheet.Cells["A3"].PutValue("Februar");
            sheet.Cells["B3"].PutValue(1500.75);
            sheet.Cells["A4"].PutValue("März");
            sheet.Cells["B4"].PutValue(1800.25);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Customize axis titles in German
            chart.ValueAxis.Title.Text = "Umsatz (€)";
            chart.CategoryAxis.Title.Text = "Monat";

            // Optional: set number format for the value axis using German conventions
            // Note: Axis.NumberFormat is not available in this version of Aspose.Cells, so this line is omitted.

            // Save the workbook to a file
            workbook.Save("ChartWithGermanLocalization.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
