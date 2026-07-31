// Title: C# – Toggle Excel Calculation Mode (Automatic ↔ Manual) Based on Workbook Size with Aspose.Cells
// Description: A reusable C# routine that loads an Excel workbook using Aspose.Cells, checks the file size, and switches FormulaSettings.CalculationMode to Manual for files larger than a specified threshold or to Automatic for smaller files, then saves the workbook.
// Keywords: Aspose.Cells | C# | Excel calculation mode | Automatic calculation | Manual calculation | file size threshold | FormulaSettings | performance optimization | toggle calculation mode | batch Excel processing
// Common Searches: Aspose.Cells set calculation mode by file size | C# toggle Excel automatic/manual calculation | change Excel calculation mode programmatically | performance tip for large Excel workbooks Aspose | how to use FormulaSettings.CalculationMode in .NET | adjust Excel calculation mode before bulk updates
// Developer Intent: Automatically apply Manual calculation to large workbooks and Automatic calculation to smaller ones to balance performance and accuracy.
// Use Cases: Speed up bulk data imports on >50 MB workbooks by switching to manual calculation before updates. | Ensure small workbooks recalculate instantly after loading, preserving expected results. | Integrate the toggler into an automated pipeline that processes multiple Excel files of varying sizes. | Provide a safety net for server‑side Excel processing where memory consumption must be controlled.
// AI Prompts: Write unit tests for CalculationModeToggler.ToggleCalculationMode covering both threshold scenarios. | Create a version of the toggler that logs the previous calculation mode and returns it to the caller. | Show how to extend the method to accept a custom CalcModeType parameter instead of hard‑coding Automatic and Manual. | Generate documentation comments for the ToggleCalculationMode method following the XML doc standard.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacro
{
    // A reusable C# routine that loads an Excel workbook using Aspose.Cells, checks the file size, and switches FormulaSettings.CalculationMode to Manual for files larger than a specified threshold or to Automatic for smaller files, then saves the workbook.
    public static class CalculationModeToggler
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sizeThresholdBytes">Size threshold in bytes.</param>
        public static void ToggleCalculationMode(string filePath, long sizeThresholdBytes)
        {
            // Ensure the file exists
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The specified workbook file was not found.", filePath);

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);

            // Determine current file size
            long fileSize = new FileInfo(filePath).Length;

            // Access formula settings
            FormulaSettings formulaSettings = workbook.Settings.FormulaSettings;

            // Toggle calculation mode based on size
            if (fileSize > sizeThresholdBytes)
            {
                // Large file -> Manual calculation to improve performance
                formulaSettings.CalculationMode = CalcModeType.Manual;
            }
            else
            {
                // Small file -> Automatic calculation
                formulaSettings.CalculationMode = CalcModeType.Automatic;
            }

            // Save the workbook (lifecycle rule: save)
            // Overwrite the original file
            workbook.Save(filePath);
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: AsposeCellsMacro <ExcelFilePath> <SizeThresholdBytes>");
                    return;
                }

                string filePath = args[0];
                if (!long.TryParse(args[1], out long sizeThresholdBytes))
                {
                    Console.WriteLine("Invalid size threshold. Please provide a numeric value.");
                    return;
                }

                // Call the toggler
                CalculationModeToggler.ToggleCalculationMode(filePath, sizeThresholdBytes);
                Console.WriteLine("Calculation mode toggled successfully.");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
