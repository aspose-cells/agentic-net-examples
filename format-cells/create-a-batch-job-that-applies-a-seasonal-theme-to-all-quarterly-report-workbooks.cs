using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace SeasonalThemeBatchJob
{
    class Program
    {
        // Entry point of the batch job
        static async Task Main(string[] args)
        {
            try
            {
                // Path to the template workbook that contains the desired seasonal theme
                string templatePath = @"C:\Templates\SeasonalThemeTemplate.xlsx";

                // Folder containing all quarterly report workbooks to be processed
                string reportsFolder = @"C:\QuarterlyReports";

                // Apply the seasonal theme to each workbook in the folder
                await ApplySeasonalThemeAsync(templatePath, reportsFolder);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Applies the theme from the template workbook to all .xlsx files in the specified folder
        private static async Task ApplySeasonalThemeAsync(string templatePath, string reportsFolder)
        {
            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.Error.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Verify that the reports folder exists
            if (!Directory.Exists(reportsFolder))
            {
                Console.Error.WriteLine($"Reports folder not found: {reportsFolder}");
                return;
            }

            try
            {
                // Load the template workbook (theme is automatically loaded with the file)
                using (Workbook templateWorkbook = new Workbook(templatePath))
                {
                    // Get all Excel files in the target folder (non-recursive for simplicity)
                    string[] reportFiles = Directory.GetFiles(reportsFolder, "*.xlsx");

                    foreach (string reportFile in reportFiles)
                    {
                        try
                        {
                            // Load the current quarterly report workbook
                            using (Workbook reportWorkbook = new Workbook(reportFile))
                            {
                                // Copy the theme from the template workbook to the report workbook
                                reportWorkbook.CopyTheme(templateWorkbook);

                                // Save the modified workbook, overwriting the original file
                                reportWorkbook.Save(reportFile);
                            }
                        }
                        catch (Exception fileEx)
                        {
                            Console.Error.WriteLine($"Failed to process '{reportFile}': {fileEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbooks: {ex.Message}");
            }

            await Task.CompletedTask; // Placeholder to satisfy async signature
        }
    }
}