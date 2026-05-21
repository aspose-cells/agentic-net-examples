using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class RemoveTempVbaModules
    {
        // Removes a VBA module named "Temp" from each workbook in the collection and saves the changes.
        public static void Run(IEnumerable<string> workbookPaths)
        {
            foreach (string path in workbookPaths)
            {
                try
                {
                    // Verify the file exists to avoid FileNotFoundException
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        continue;
                    }

                    // Load the workbook from file
                    using (Workbook workbook = new Workbook(path))
                    {
                        // Access the VBA project (may be empty)
                        VbaProject vbaProject = workbook.VbaProject;

                        // Attempt to remove the module named "Temp" if it exists
                        // The Remove method does not throw if the name is not found
                        vbaProject?.Modules?.Remove("Temp");

                        // Save the modified workbook (overwrites the original file)
                        workbook.Save(path, SaveFormat.Xlsx);
                    }

                    Console.WriteLine($"Processed workbook: {path}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }
        }
    }

    public class Program
    {
        // Entry point required for the application
        public static void Main(string[] args)
        {
            try
            {
                if (args == null || args.Length == 0)
                {
                    Console.WriteLine("Please provide one or more workbook file paths as arguments.");
                    return;
                }

                // Pass the provided file paths to the processing method
                RemoveTempVbaModules.Run(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}