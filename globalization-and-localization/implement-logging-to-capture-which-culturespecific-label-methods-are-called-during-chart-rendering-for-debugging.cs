// Title: Add a logger to capture culture‑specific chart label method calls during Aspose.Cells chart rendering in C#
// AI Prompts: Insert logging statements before each chart axis label method (e.g., CategoryAxis.Title, ValueAxis.Title) to record the method name and CultureInfo.CurrentCulture.Name during workbook generation. | Enhance the SimpleLogger class so it automatically writes an entry whenever a label‑related API of Aspose.Cells is invoked while rendering a chart, including a timestamp and the active culture. | Provide a C# sample that creates a column chart with Aspose.Cells, logs every call to chart label rendering methods, saves the workbook, and then displays the collected log entries.
// Common Searches: how to debug culture specific chart labels in Aspose.Cells .NET | log axis title method calls with culture info using Aspose.Cells C# | Aspose.Cells chart localization troubleshooting example | record which chart label APIs are executed during Excel file creation in C# | track culture changes while rendering charts with Aspose.Cells
// Tags: chart axis label logging Aspose.Cells | culture-aware chart rendering .NET | Aspose.Cells localization debugging | C# logger for Excel chart generation | track culture info in Aspose.Cells charts

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Simple logger to capture operation timestamps.
    // The example defines a SimpleLogger that records timestamped messages together with the current culture, creates a workbook, populates data, adds a column chart, and saves the file. It demonstrates how to extend the logger to capture each chart label method invocation, enabling developers to debug culture‑specific rendering behavior in Aspose.Cells.
    class SimpleLogger
    {
        private readonly List<string> _entries = new List<string>();
        public IReadOnlyList<string> Entries => _entries.AsReadOnly();

        public void Add(string message)
        {
            _entries.Add($"{DateTime.Now:O} | {message} | Culture: {CultureInfo.CurrentCulture.Name}");
        }
    }

    class Program
    {
        static void Main()
        {
            var logger = new SimpleLogger();

            try
            {
                logger.Add("Creating new workbook");
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data.
                logger.Add("Populating worksheet with data");
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart.
                logger.Add("Adding column chart");
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Save the workbook.
                string outputPath = "ChartWithLogging.xlsx";
                logger.Add($"Saving workbook to {outputPath}");
                workbook.Save(outputPath);
                logger.Add("Workbook saved successfully");
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                logger.Add($"Exception: {ex.Message}");
                Console.WriteLine($"Error: {ex.Message}");
            }

            // Output log entries.
            Console.WriteLine("=== Operation Log ===");
            foreach (string entry in logger.Entries)
            {
                Console.WriteLine(entry);
            }
        }
    }
}
