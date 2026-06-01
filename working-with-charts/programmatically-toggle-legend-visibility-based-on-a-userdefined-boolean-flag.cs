using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendToggle
{
    public class Program
    {
        // Entry point
        public static void Main(string[] args)
        {
            // Example flag – in real scenario this could come from user input, config, etc.
            bool showLegend = GetUserDefinedFlag();

            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Toggle legend visibility based on the flag (Chart.ShowLegend property rule)
            chart.ShowLegend = showLegend;
            Console.WriteLine($"Legend visibility set to: {chart.ShowLegend}");

            // Save the workbook (save rule)
            workbook.Save("ChartWithLegendToggle.xlsx");
        }

        // Mock method to obtain a user‑defined boolean flag
        private static bool GetUserDefinedFlag()
        {
            // For demonstration, toggle based on console input.
            Console.Write("Display chart legend? (y/n): ");
            string input = Console.ReadLine()?.Trim().ToLowerInvariant();
            return input == "y" || input == "yes";
        }
    }
}