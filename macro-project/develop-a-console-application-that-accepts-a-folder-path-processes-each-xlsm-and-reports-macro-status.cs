using System;
using System.IO;
using Aspose.Cells;

namespace MacroStatusReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide a folder path as the first argument.");
                return;
            }

            string folderPath = args[0];

            // Check if the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Get all .xlsm files in the folder (non‑recursive)
            string[] xlsmFiles = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            if (xlsmFiles.Length == 0)
            {
                Console.WriteLine("No macro‑enabled Excel files (*.xlsm) were found in the specified folder.");
                return;
            }

            Console.WriteLine($"Processing {xlsmFiles.Length} file(s) in \"{folderPath}\":");

            foreach (string filePath in xlsmFiles)
            {
                try
                {
                    // Load the workbook (uses the Workbook(string) constructor rule)
                    Workbook workbook = new Workbook(filePath);

                    // Check if the workbook contains macros (uses the HasMacro property rule)
                    bool hasMacro = workbook.HasMacro;

                    // Report the result
                    Console.WriteLine($"{Path.GetFileName(filePath)} : HasMacro = {hasMacro}");
                }
                catch (Exception ex)
                {
                    // Report any errors encountered while processing the file
                    Console.WriteLine($"{Path.GetFileName(filePath)} : Error - {ex.Message}");
                }
            }
        }
    }
}