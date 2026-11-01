// Title: Check each VBA module for a Sub Main procedure before saving an Excel workbook with Aspose.Cells in C#
// Common Searches: how to verify that every VBA module contains Sub Main using Aspose.Cells C# | Aspose.Cells C# throw error when VBA module missing Sub Main before workbook save | validate VBA macro entry point in .xlsx file with Aspose.Cells .NET | C# code to ensure VBA project modules have Sub Main before saving workbook | Aspose.Cells workbook validation for required Sub Main in VBA modules
// Tags: Aspose.Cells VBA module validation | C# check Sub Main in Excel macros | Excel workbook macro entry point verification | Aspose.Cells workbook save pre‑validation | InvalidOperationException missing Sub Main Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace WorkbookValidation
{
    // The example loads an Excel workbook, iterates through its VBA project modules, and throws an InvalidOperationException if any module lacks a "Sub Main" procedure; otherwise it saves the workbook to the specified destination.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source and destination workbooks
            string sourcePath = @"C:\Temp\SourceWorkbook.xlsx";
            string destinationPath = @"C:\Temp\ValidatedWorkbook.xlsx";

            try
            {
                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Perform VBA validation
                ValidateModulesContainSubMain(workbook);

                // Ensure destination directory exists
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the validated workbook
                workbook.Save(destinationPath);
                Console.WriteLine($"Workbook saved successfully to {destinationPath}");
            }
            catch (InvalidOperationException ex)
            {
                // Validation failure
                Console.WriteLine($"Validation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="workbook">The Aspose.Cells workbook to validate.</param>
        private static void ValidateModulesContainSubMain(Workbook workbook)
        {
            // No VBA project means nothing to validate
            if (workbook.VbaProject == null)
                return;

            // Iterate through all VBA modules
            foreach (VbaModule module in workbook.VbaProject.Modules)
            {
                string code = module.Codes ?? string.Empty;

                // Check for "Sub Main"
                if (!code.Contains("Sub Main", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"VBA module \"{module.Name}\" does not contain a required \"Sub Main\" entry point.");
                }
            }
        }
    }
}
