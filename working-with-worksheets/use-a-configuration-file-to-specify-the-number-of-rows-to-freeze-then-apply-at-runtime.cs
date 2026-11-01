// Title: Freeze a configurable number of top rows in an Excel worksheet using Aspose.Cells and a JSON config file (C#)
// AI Prompts: Generate C# code that reads a JSON configuration file to obtain a row count and applies Worksheet.FreezePanes to freeze that many top rows with Aspose.Cells. | Demonstrate how to create a Config class and a LoadConfig method for dynamic row freezing in an Aspose.Cells workbook. | Provide a sample config.json and show the complete workflow of saving the workbook after applying the freeze panes based on the config value.
// Common Searches: how to freeze rows in Aspose.Cells using a value from a json file | c# read json config to set freeze panes in excel with Aspose.Cells | dynamic number of frozen rows Aspose.Cells .NET example | using Worksheet.FreezePanes overload with runtime parameters in C# | load json settings to control Excel freeze panes Aspose.Cells
// Tags: Aspose.Cells FreezePanes runtime configuration | C# load JSON settings for Excel freeze rows | dynamic top row freeze Aspose.Cells workbook | Worksheet.FreezePanes overload usage | configurable frozen rows Excel .NET

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace FreezeRowsExample
{
    // Simple configuration class matching the JSON structure
    // The example reads a config.json file to get a FreezeRows value, creates a workbook, optionally fills sample data, and calls Worksheet.FreezePanes with that value to freeze the specified number of top rows before saving the file.
    public class Config
    {
        public int FreezeRows { get; set; } = 0; // default to no frozen rows
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the configuration file (e.g., config.json)
            string configPath = "config.json";

            // Load configuration
            Config config = LoadConfig(configPath);

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data (optional, just for demonstration)
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Apply freeze panes based on the configuration
                // Use the 4‑argument overload for compatibility with all Aspose.Cells versions
                // FreezeRows rows from the top, no columns are frozen
                sheet.FreezePanes(config.FreezeRows, 0, config.FreezeRows, 0);

                // Save the workbook to a file
                string outputPath = "FrozenRowsWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}' with {config.FreezeRows} frozen row(s).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while creating or saving the workbook: {ex.Message}");
            }
        }

        // Helper method to load configuration from a JSON file
        private static Config LoadConfig(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine($"Configuration file '{path}' not found. Using default settings.");
                return new Config();
            }

            try
            {
                string json = File.ReadAllText(path);
                Config? cfg = JsonSerializer.Deserialize<Config>(json);
                if (cfg == null)
                {
                    Console.WriteLine("Configuration file is empty or invalid. Using default settings.");
                    return new Config();
                }
                return cfg;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading configuration: {ex.Message}. Using default settings.");
                return new Config();
            }
        }
    }
}
