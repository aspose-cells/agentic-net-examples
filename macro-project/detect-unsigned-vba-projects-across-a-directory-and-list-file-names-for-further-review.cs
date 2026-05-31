using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – change as needed
            string folderPath = @"C:\ExcelFiles";

            // Collect files with unsigned VBA projects
            List<string> unsignedVbaFiles = new List<string>();

            // Supported Excel extensions (including macro-enabled)
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xls2003", ".xls2007" };

            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                // Skip files that are not Excel based on extension
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                try
                {
                    // Load the workbook (read‑only is sufficient for inspection)
                    Workbook workbook = new Workbook(filePath);

                    // If the workbook contains VBA/macros
                    if (workbook.HasMacro)
                    {
                        // Check if the VBA project is signed
                        bool isSigned = workbook.VbaProject.IsSigned;

                        // If not signed, add to the result list
                        if (!isSigned)
                        {
                            unsignedVbaFiles.Add(filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log or ignore files that cannot be processed as Excel workbooks
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Output the list of files with unsigned VBA projects
            Console.WriteLine("Files with unsigned VBA projects:");
            foreach (string file in unsignedVbaFiles)
            {
                Console.WriteLine(file);
            }
        }
    }
}