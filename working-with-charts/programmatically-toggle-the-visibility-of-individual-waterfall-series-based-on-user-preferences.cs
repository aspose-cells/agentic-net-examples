// Title: How to programmatically hide or show individual series in an Aspose.Cells waterfall chart (C#)
// Common Searches: C# Aspose.Cells hide specific series in a waterfall chart | toggle visibility of waterfall chart series using Aspose.Cells API | set IsFiltered property for chart series in Aspose.Cells example | user input driven series show hide in Excel waterfall chart with Aspose.Cells | Aspose.Cells programmatic control of series visibility in waterfall chart
// Tags: Aspose.Cells waterfall chart series visibility | C# chart series IsFiltered property | programmatic series hide show Excel | dynamic chart series filtering Aspose.Cells | Excel waterfall chart series control C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallSeriesVisibilityDemo
{
    // The example creates a workbook with sample data, adds a waterfall chart containing three series, reads series indices from console input, toggles each selected series' IsFiltered property to hide or show it, and saves the workbook as WaterfallSeriesVisibility.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and fill it with sample data for a waterfall chart
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Sample categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Revenue");
                sheet.Cells["A4"].PutValue("Cost");
                sheet.Cells["A5"].PutValue("Profit");
                sheet.Cells["A6"].PutValue("End");

                // Sample values (waterfall works with positive/negative numbers)
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(0);    // Start (blank)
                sheet.Cells["B3"].PutValue(120);  // Revenue
                sheet.Cells["B4"].PutValue(-40);  // Cost
                sheet.Cells["B5"].PutValue(80);   // Profit (calculated automatically)
                sheet.Cells["B6"].PutValue(0);    // End (blank)

                // -----------------------------------------------------------------
                // 2. Add a Waterfall chart
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Waterfall Demo";

                // Add the main series (the only series needed for a basic waterfall)
                int seriesIdx = chart.NSeries.Add("B2:B6", true);
                // Set category (X‑axis) data for the whole NSeries collection
                chart.NSeries.CategoryData = "A2:A6";

                // For demonstration, add two more dummy series that will be toggled.
                // These series use the same data range but are displayed as separate columns.
                // In a real waterfall you normally have a single series, but Aspose.Cells
                // allows multiple series for the purpose of this demo.
                int extraSeries1 = chart.NSeries.Add("B2:B6", true);
                chart.NSeries[extraSeries1].Name = "Extra Series 1";

                int extraSeries2 = chart.NSeries.Add("B2:B6", true);
                chart.NSeries[extraSeries2].Name = "Extra Series 2";

                // -----------------------------------------------------------------
                // 3. Display current visibility status of each series
                // -----------------------------------------------------------------
                Console.WriteLine("Current series visibility (false = visible, true = hidden):");
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Console.WriteLine($"  Series {i}: IsFiltered = {chart.NSeries[i].IsFiltered}");
                }

                // -----------------------------------------------------------------
                // 4. Ask user which series to toggle
                // -----------------------------------------------------------------
                Console.WriteLine();
                Console.WriteLine("Enter the indices of the series you want to toggle visibility.");
                Console.WriteLine("Separate multiple indices with commas (e.g., 0,2).");
                Console.Write("Your input: ");
                string input = Console.ReadLine();

                if (!string.IsNullOrWhiteSpace(input))
                {
                    string[] parts = input.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    List<int> indices = new List<int>();

                    foreach (string part in parts)
                    {
                        if (int.TryParse(part.Trim(), out int idx))
                        {
                            if (idx >= 0 && idx < chart.NSeries.Count)
                                indices.Add(idx);
                            else
                                Console.WriteLine($"Index {idx} is out of range and will be ignored.");
                        }
                        else
                        {
                            Console.WriteLine($"Unable to parse '{part}' as an integer and will be ignored.");
                        }
                    }

                    // -----------------------------------------------------------------
                    // 5. Toggle visibility for the selected series
                    // -----------------------------------------------------------------
                    foreach (int idx in indices)
                    {
                        Series s = chart.NSeries[idx];
                        // The IsFiltered property hides the series when set to true.
                        // We toggle it: if it was hidden, make it visible; otherwise hide it.
                        s.IsFiltered = !s.IsFiltered;
                        Console.WriteLine($"Series {idx} visibility toggled. New IsFiltered = {s.IsFiltered}");
                    }
                }
                else
                {
                    Console.WriteLine("No input received. No changes will be made.");
                }

                // -----------------------------------------------------------------
                // 6. Save the workbook
                // -----------------------------------------------------------------
                string outFile = "WaterfallSeriesVisibility.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine();
                Console.WriteLine($"Workbook saved as '{outFile}'. Open it to see the result.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
