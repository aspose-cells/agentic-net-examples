// Title: Set German locale for Excel chart axis labels using Aspose.Cells C# example
// Common Searches: how to display Excel chart axis labels in German with Aspose.Cells C# | Aspose.Cells chart globalization to German culture example | C# code to apply de-DE culture to chart axis labels in an Excel workbook | using Aspose.Cells workbook Settings.CultureInfo for German chart localization
// Tags: Aspose.Cells workbook culture German | chart axis label localization Aspose.Cells | C# Excel chart globalization de-DE | apply German locale to chart series Aspose.Cells | ChartGlobalizationSettings example C#

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, sets its Settings.CultureInfo to "de-DE", fills sample data, adds a column chart, assigns series and category labels, and saves the file as ChartGerman.xlsx. It demonstrates workbook‑level culture affecting chart axis labels but does not instantiate a ChartGlobalizationSettings object.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set workbook culture to German (affects number/date formatting)
                workbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the chart
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Jan");
                cells["A3"].PutValue("Feb");
                cells["A4"].PutValue("Mar");
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series and category (axis) labels
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Save the workbook
                string outputPath = "ChartGerman.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
