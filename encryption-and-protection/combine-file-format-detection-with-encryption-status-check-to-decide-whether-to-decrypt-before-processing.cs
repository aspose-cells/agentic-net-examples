// Title: Detect Excel file format and decrypt encrypted workbook with Aspose.Cells LoadOptions in C#
// Common Searches: C# load Excel file with unknown format and decrypt if password protected using Aspose.Cells | How to automatically detect workbook format and handle encrypted Excel files in Aspose.Cells | Aspose.Cells LoadOptions Auto format detection with password fallback in C# | Process encrypted Excel workbook after format detection with Aspose.Cells | C# example for catching CellsException to load password protected Excel file
// Tags: loadoptions automatic format detection Aspose.Cells | decrypt password protected workbook C# | conditional workbook loading based on encryption Aspose.Cells | catch CellsException for encrypted Excel files | save processed workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The ExcelProcessor class checks the input file, uses LoadOptions with LoadFormat.Auto to detect the Excel format, attempts to open the workbook, catches a CellsException for encrypted files, and if a password is provided reloads the workbook with the password. It then makes the first worksheet visible, ensures the output directory exists, and saves the processed workbook.
    public class ExcelProcessor
    {
        /// <param name="inputPath">Path to the source Excel file.</param>
        /// <param name="outputPath">Path where the processed file will be saved.</param>
        /// <param name="password">Password for decryption (null or empty if not needed).</param>
        public void ProcessFile(string inputPath, string outputPath, string? password = null)
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load options with automatic format detection.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            Workbook workbook = null;

            try
            {
                // Attempt to load without a password.
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (CellsException ex)
            {
                // If the workbook is encrypted and a password is supplied, retry with the password.
                if (!string.IsNullOrEmpty(password) && ex.Message.Contains("encrypted", StringComparison.OrdinalIgnoreCase))
                {
                    loadOptions.Password = password;
                    workbook = new Workbook(inputPath, loadOptions);
                }
                else
                {
                    // Re‑throw if it's not an encryption issue or no password was provided.
                    throw;
                }
            }

            // Example processing: ensure the first worksheet is visible.
            if (workbook?.Worksheets?.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.IsVisible = true;
            }

            // Ensure the output directory exists.
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the processed workbook.
            workbook?.Save(outputPath);
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            // Adjust these paths as needed for testing.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string? password = null; // Set password if the file is encrypted.

            try
            {
                var processor = new ExcelProcessor();
                processor.ProcessFile(inputPath, outputPath, password);
                Console.WriteLine("Excel file processed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
