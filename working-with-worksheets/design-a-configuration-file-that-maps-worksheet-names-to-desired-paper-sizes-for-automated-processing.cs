using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPaperSizeConfig
{
    /// <summary>
    /// Demonstrates how to apply worksheet‑specific paper sizes based on a simple
    /// configuration file. The config file contains one mapping per line in the form:
    /// SheetName=PaperSizeEnum
    /// Example:
    /// Sheet1=PaperA4
    /// Report=PaperLetter
    /// </summary>
    public class PaperSizeConfigurator
    {
        /// <summary>
        /// Loads the mapping from the configuration file.
        /// </summary>
        /// <param name="configPath">Path to the config file.</param>
        /// <returns>Dictionary where key = worksheet name, value = PaperSizeType enum name.</returns>
        private static Dictionary<string, string> LoadConfig(string configPath)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                foreach (var line in File.ReadAllLines(configPath))
                {
                    // Skip empty lines and comments
                    if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#"))
                        continue;

                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length != 2)
                        continue; // malformed line – ignore

                    var sheetName = parts[0].Trim();
                    var paperSize = parts[1].Trim();

                    if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(paperSize))
                        map[sheetName] = paperSize;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading config file \"{configPath}\": {ex.Message}");
            }

            return map;
        }

        /// <summary>
        /// Applies the paper size settings to the workbook according to the mapping.
        /// </summary>
        /// <param name="workbook">The workbook to modify.</param>
        /// <param name="sheetPaperMap">Mapping of sheet name to PaperSizeType name.</param>
        private static void ApplyPaperSizes(Workbook workbook, Dictionary<string, string> sheetPaperMap)
        {
            foreach (var kvp in sheetPaperMap)
            {
                string sheetName = kvp.Key;
                string paperSizeName = kvp.Value;

                // Verify the worksheet exists
                Worksheet sheet = null;
                try
                {
                    sheet = workbook.Worksheets[sheetName];
                }
                catch (ArgumentException)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found. Skipping.");
                    continue;
                }

                // Convert the string to the corresponding enum value
                if (Enum.TryParse<PaperSizeType>(paperSizeName, out var paperSize))
                {
                    sheet.PageSetup.PaperSize = paperSize;
                    Console.WriteLine($"Set paper size of \"{sheetName}\" to {paperSize}.");
                }
                else
                {
                    Console.WriteLine($"Invalid paper size \"{paperSizeName}\" for sheet \"{sheetName}\". Skipping.");
                }
            }
        }

        /// <summary>
        /// Entry point – loads a workbook, reads the config, applies settings, and saves the result.
        /// </summary>
        public static void Run()
        {
            // Paths – adjust as needed
            string inputWorkbookPath = "input.xlsx";
            string outputWorkbookPath = "output.xlsx";
            string configFilePath = "paperSizeConfig.txt";

            // Verify input files exist
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Input workbook \"{inputWorkbookPath}\" not found.");
                return;
            }

            if (!File.Exists(configFilePath))
            {
                Console.WriteLine($"Configuration file \"{configFilePath}\" not found.");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook (lifecycle rule: load)
                workbook = new Workbook(inputWorkbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{inputWorkbookPath}\": {ex.Message}");
                return;
            }

            // Load configuration
            var sheetPaperMap = LoadConfig(configFilePath);

            // Apply the paper size settings
            ApplyPaperSizes(workbook, sheetPaperMap);

            try
            {
                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputWorkbookPath}\" with updated paper sizes.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook \"{outputWorkbookPath}\": {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            PaperSizeConfigurator.Run();
        }
    }
}