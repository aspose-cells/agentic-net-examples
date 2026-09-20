// Title: Configure worksheet-specific paper sizes from a JSON file with Aspose.Cells for .NET
// AI Prompts: Read a JSON mapping of worksheet names to paper size strings, convert each string to the Aspose.Cells PaperSizeType enum, and assign it to the worksheet's PageSetup.PaperSize in C#. | Add validation that logs a warning when a JSON entry contains an unsupported paper size before attempting to set the PageSetup.PaperSize. | Generate a sample worksheetPaperSizes.json that maps sheet names to standard sizes like A4, Letter, and Legal for use with Aspose.Cells.
// Common Searches: Aspose.Cells set different paper size for each worksheet using JSON configuration | C# load paper size mapping from json and apply to Excel sheets with Aspose.Cells | How to map worksheet names to PaperSizeType enum in Aspose.Cells .NET | Error handling unknown paper size values when using Aspose.Cells PageSetup
// Tags: worksheet page setup paper size Aspose.Cells | JSON to PaperSizeType mapping C# | apply per-sheet paper size Aspose.Cells | handle unsupported paper size enum Aspose.Cells | load workbook and configure page setup programmatically

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace WorksheetPaperSizeConfigurator
{
    // Represents the mapping between worksheet names and desired paper sizes.
    // The example reads a JSON file that maps worksheet names to paper size identifiers, loads an Excel workbook with Aspose.Cells, iterates through each worksheet, parses the corresponding PaperSizeType enum (e.g., PaperA4) and assigns it to the worksheet's PageSetup.PaperSize, then saves the updated workbook while handling missing files and unknown size values.
    public class PaperSizeConfig : Dictionary<string, string> { }

    class Program
    {
        static void Main()
        {
            // Paths to the workbook and configuration files.
            const string workbookPath = "InputWorkbook.xlsx";
            const string configPath = "worksheetPaperSizes.json";
            const string outputPath = "OutputWorkbook.xlsx";

            try
            {
                // Verify that the input workbook exists.
                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

                // Load the configuration.
                PaperSizeConfig config = LoadConfig(configPath);

                // Load the workbook.
                Workbook workbook;
                try
                {
                    workbook = new Workbook(workbookPath);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to load workbook '{workbookPath}'.", ex);
                }

                // Apply paper size settings to each worksheet based on the configuration.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (config.TryGetValue(sheet.Name, out string paperSizeName))
                    {
                        // Build the enum name (e.g., "PaperA4").
                        string enumName = "Paper" + paperSizeName;

                        // Try to parse the enum value using PaperSizeType (the correct enum in Aspose.Cells).
                        if (Enum.TryParse<PaperSizeType>(enumName, true, out var paperSizeEnum))
                        {
                            sheet.PageSetup.PaperSize = paperSizeEnum;
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Unknown paper size '{paperSizeName}' for worksheet '{sheet.Name}'.");
                        }
                    }
                }

                // Save the modified workbook.
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to save workbook to '{outputPath}'.", ex);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Loads the JSON configuration file into a PaperSizeConfig dictionary.
        private static PaperSizeConfig LoadConfig(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Configuration file not found: {path}");

            try
            {
                string json = File.ReadAllText(path);
                var config = JsonSerializer.Deserialize<PaperSizeConfig>(json);
                return config ?? new PaperSizeConfig();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to read or parse configuration file '{path}'.", ex);
            }
        }
    }
}
