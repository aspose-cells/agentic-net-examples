using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationSave
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains localized content,
            // formatting, comments, and cell styles.
            string sourcePath = "LocalizedInput.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source file '{sourcePath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook. The constructor with a file path
                // utilizes the built‑in load rule.
                Workbook workbook = new Workbook(sourcePath);

                // Save the workbook as XLSX while preserving all original
                // formatting, comments, and styles.
                string outputPath = "LocalizedOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions (e.g., loading/saving errors) and display a message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}