// Title: Create a Column Chart with Spanish Month Legends using Aspose.Cells for .NET
// Description: Shows how to build an Excel workbook, fill month numbers and values, add a column chart, set short Spanish month abbreviations as series names so they appear in the legend, print the legend entries for verification, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# chart legend localization | Spanish month names | Excel column chart series name | .NET Aspose.Cells example | custom chart legend | locale chart legend | Aspose.Cells chart series | Excel localization | Spanish sales chart
// Common Searches: Aspose.Cells chart legend in Spanish | set custom series names for Excel chart using C# | display month abbreviations in chart legend .NET | verify Aspose.Cells legend entries programmatically | localize Excel chart legends with Aspose.Cells
// Developer Intent: Generate an Excel file with a column chart whose legend shows short Spanish month names.
// Use Cases: Monthly sales report for Spanish‑speaking stakeholders with a localized legend. | Chart where each month is a separate series to allow per‑month styling or data updates. | Automated validation that chart legend entries match expected Spanish month abbreviations before distribution.
// AI Prompts: Write C# code with Aspose.Cells to create a column chart, add twelve series, and name each series using short Spanish month abbreviations, then output the legend entries. | Show how to programmatically compare Aspose.Cells chart series names against a predefined list of Spanish month names and report mismatches. | Explain how to modify the example to use full Spanish month names instead of abbreviations while keeping the legend correctly localized.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build an Excel workbook, fill month numbers and values, add a column chart, set short Spanish month abbreviations as series names so they appear in the legend, print the legend entries for verification, and save the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (12 months, one value each)
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 0; i < 12; i++)
                {
                    // Use month number as placeholder; legend will use Spanish names
                    sheet.Cells[i + 1, 0].PutValue(i + 1);
                    sheet.Cells[i + 1, 1].PutValue(10 * (i + 1)); // arbitrary values
                }

                // Create a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Ventas Mensuales";

                // Spanish short month names
                string[] spanishMonths = new string[]
                {
                    "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                    "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
                };

                // Add each month as a separate series so that the legend shows month names
                for (int i = 0; i < 12; i++)
                {
                    // Each series references a single cell in column B (value)
                    int seriesIndex = chart.NSeries.Add($"B{i + 2}", false);
                    // Set the series name to the Spanish month name
                    chart.NSeries[seriesIndex].Name = spanishMonths[i];
                }

                // Verify that the series (legend) names are set correctly
                Console.WriteLine("Legend entries (Spanish month names):");
                foreach (Series series in chart.NSeries)
                {
                    Console.WriteLine($"- {series.Name}");
                }

                // Save the workbook
                string outputPath = "ChartWithSpanishMonthLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
