// Title: Apply German CultureInfo to an Aspose.Cells chart to localize axis labels in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, sets Workbook.Settings.CultureInfo to "de-DE", adds a column chart, and saves the file. | Show how to use ChartGlobalizationSettings to format chart axis numbers and dates for the German locale in Aspose.Cells. | Provide a step‑by‑step example of localizing axis label language on a specific chart using Aspose.Cells for .NET.
// Common Searches: how to set german locale for chart axis labels using Aspose.Cells C# | Aspose.Cells chart globalization settings de-DE example | C# code to localize Excel chart axis to German with Aspose | apply cultureinfo to specific chart in Aspose.Cells .NET
// Tags: Aspose.Cells set chart cultureinfo | C# chart localization de-DE | ChartGlobalizationSettings German formatting | Excel chart axis localization Aspose | Workbook.Settings.CultureInfo chart example

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Creates a workbook, sets its CultureInfo to German (de-DE), adds sample data, inserts a column chart, and saves the file as ChartGerman.xlsx, demonstrating how the German locale automatically localizes chart axis labels.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Apply German culture settings to the entire workbook (affects charts)
                workbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B1:B3", true);
                chart.NSeries.CategoryData = "A1:A3";

                // Save the workbook
                string outputPath = "ChartGerman.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
