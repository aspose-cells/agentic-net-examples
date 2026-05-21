using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsChartExclusionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that may contain charts
            string sourcePath = "input_with_charts.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: File not found – \"{sourcePath}\"");
                return;
            }

            try
            {
                // Load the workbook (charts will be removed manually after loading)
                Workbook workbook = new Workbook(sourcePath);

                // Remove all charts from each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Charts.Clear();
                }

                // Verify that each worksheet now contains zero charts
                bool allSheetsHaveNoCharts = true;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int chartCount = sheet.Charts.Count;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" has {chartCount} chart(s).");
                    if (chartCount != 0)
                    {
                        allSheetsHaveNoCharts = false;
                    }
                }

                Console.WriteLine(allSheetsHaveNoCharts
                    ? "All worksheets contain zero charts."
                    : "Some worksheets still contain charts.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}