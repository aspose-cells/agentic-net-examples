using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapCreator
{
    // Console application that asks for an XML schema (XSD) file path,
    // creates a new Excel workbook with an XML map based on the schema,
    // and saves the workbook to a user‑specified location.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Prompt the user to enter the XML schema file path
                string schemaPath = PromptForExistingFile("Enter full path to the XML Schema (XSD) file:");
                if (string.IsNullOrEmpty(schemaPath))
                {
                    Console.WriteLine("No schema file provided. Exiting.");
                    return;
                }

                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Add the selected XML schema as a map to the workbook (lifecycle: modify)
                int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Set a friendly name for the map (file name without extension)
                xmlMap.Name = Path.GetFileNameWithoutExtension(schemaPath);
                Console.WriteLine($"XML map '{xmlMap.Name}' added successfully.");

                // Prompt the user to enter the output Excel file path
                string savePath = PromptForSavePath("Enter full path where the workbook should be saved (including .xlsx):");
                if (string.IsNullOrEmpty(savePath))
                {
                    Console.WriteLine("No save location provided. Exiting without saving.");
                    return;
                }

                // Ensure the target directory exists
                string? directory = Path.GetDirectoryName(savePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook with the new XML map (lifecycle: save)
                workbook.Save(savePath);
                Console.WriteLine($"Workbook saved to: {savePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Prompts the user for a file path and verifies that the file exists.
        private static string PromptForExistingFile(string message)
        {
            Console.WriteLine(message);
            string? path = Console.ReadLine()?.Trim('"').Trim();

            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                Console.WriteLine("File not found.");
                return string.Empty;
            }

            return path;
        }

        // Prompts the user for a save path (does not need to exist beforehand).
        private static string PromptForSavePath(string message)
        {
            Console.WriteLine(message);
            string? path = Console.ReadLine()?.Trim('"').Trim();

            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine("Invalid path.");
                return string.Empty;
            }

            // Ensure the file has an .xlsx extension
            if (!Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                path = Path.ChangeExtension(path, ".xlsx");
            }

            return path;
        }
    }
}