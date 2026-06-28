using System;
using System.IO;
using Aspose.Cells;

namespace MacroChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was supplied
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: MacroChecker <folderPath>");
                return;
            }

            string folderPath = args[0];

            // Ensure the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Retrieve all .xlsm files in the specified folder
            string[] macroFiles = Directory.GetFiles(folderPath, "*.xlsm");

            if (macroFiles.Length == 0)
            {
                Console.WriteLine("No macro-enabled Excel files (*.xlsm) found in the folder.");
                return;
            }

            // Process each file
            foreach (string filePath in macroFiles)
            {
                try
                {
                    // Load the workbook (uses Aspose.Cells load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Determine if the workbook contains macros
                    bool hasMacro = workbook.HasMacro;

                    // Output the result
                    Console.WriteLine($"{Path.GetFileName(filePath)} : HasMacro = {hasMacro}");

                    // Release resources
                    workbook.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }
        }
    }
}