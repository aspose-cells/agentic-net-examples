using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsVbaAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – use first argument or current directory if none provided
            string targetDirectory = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            // Collection to hold file paths of unsigned VBA projects
            List<string> unsignedVbaFiles = new List<string>();

            // Define file extensions that may contain VBA macros
            string[] macroExtensions = new[] { ".xlsm", ".xlsb", ".xls" };

            // Enumerate all files with the specified extensions recursively
            IEnumerable<string> files = Directory.EnumerateFiles(targetDirectory, "*.*", SearchOption.AllDirectories)
                                                .Where(f => macroExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

            foreach (string filePath in files)
            {
                try
                {
                    // Load the workbook (uses Aspose.Cells load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Check if the workbook actually contains a macro/VBA project
                    if (workbook.HasMacro)
                    {
                        // Determine whether the VBA project is signed
                        if (!workbook.VbaProject.IsSigned)
                        {
                            unsignedVbaFiles.Add(filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors (e.g., corrupted files) and continue processing
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Output the results
            Console.WriteLine("Unsigned VBA projects found:");
            foreach (string unsignedFile in unsignedVbaFiles)
            {
                Console.WriteLine(unsignedFile);
            }
        }
    }
}