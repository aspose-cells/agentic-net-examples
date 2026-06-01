using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Custom globalization settings that returns Spanish month names
    public class SpanishPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        public override string[] GetShortTextOf12Months()
        {
            // Spanish short month names
            return new string[]
            {
                "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
            };
        }
    }

    public class ChartLegendSpanishMonthsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Obtain Spanish month names from the custom globalization settings
                SpanishPivotGlobalizationSettings monthSettings = new SpanishPivotGlobalizationSettings();
                string[] spanishMonths = monthSettings.GetShortTextOf12Months();

                // Populate worksheet with month names (as categories) and sample values
                // Column A: Month names, Column B: Sample numeric values
                for (int i = 0; i < spanishMonths.Length; i++)
                {
                    // Row index starts at 1 because row 0 will be the header
                    int row = i + 1;
                    sheet.Cells[row, 0].PutValue(spanishMonths[i]);          // Category (month)
                    sheet.Cells[row, 1].PutValue((i + 1) * 10);              // Sample value
                }

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Ventas Mensuales";

                // Add each month as a separate series so that the legend shows month names
                for (int i = 0; i < spanishMonths.Length; i++)
                {
                    int row = i + 1; // Data row for the current month
                    // Add a series that references the single value cell in column B
                    chart.NSeries.Add($"B{row + 1}", false);
                    // Set the series name – this appears in the chart legend
                    chart.NSeries[i].Name = spanishMonths[i];
                }

                // Verify the legend entries by printing them to the console
                Console.WriteLine("Legend entries (should be Spanish month names):");
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Console.WriteLine($"Series {i + 1}: {chart.NSeries[i].Name}");
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "ChartLegendSpanishMonths.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved as {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ChartLegendSpanishMonthsDemo.Run();
        }
    }
}