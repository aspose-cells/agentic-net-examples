using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookThemeProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: process all Excel files in a folder
            string inputFolder = @"C:\InputWorkbooks";
            string outputFolder = @"C:\OutputWorkbooks";

            ProcessWorkbooks(inputFolder, outputFolder);
        }

        /// <summary>
        /// Loads each workbook from the input folder, checks if it has a theme,
        /// and assigns a default theme when missing before saving to the output folder.
        /// </summary>
        static void ProcessWorkbooks(string inputFolder, string outputFolder)
        {
            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Create a source workbook that contains the default theme.
            // A newly created workbook automatically has the built‑in default theme.
            using (Workbook defaultThemeWorkbook = new Workbook())
            {
                // Iterate through all supported Excel files in the input folder
                foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
                {
                    // Filter by common Excel extensions
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb")
                        continue;

                    // Load the workbook
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Detect missing theme: Theme property returns null or empty string when no theme is set
                        if (string.IsNullOrEmpty(workbook.Theme))
                        {
                            // Assign the default theme by copying it from the source workbook
                            workbook.CopyTheme(defaultThemeWorkbook);
                        }

                        // Save the processed workbook to the output folder (preserving original file name)
                        string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                        workbook.Save(outputPath);
                    }
                }
            }
        }
    }
}