// Title: Find unsigned VBA projects in multiple Excel workbooks using Aspose.Cells for .NET
// Description: Iterates over a collection of workbook paths, loads each file with Aspose.Cells, checks Workbook.HasMacro and VbaProject.IsSigned, and writes the names of workbooks with unsigned VBA projects to the console while handling missing files and runtime errors.
// Keywords: Aspose.Cells unsigned VBA detection | C# macro project signing check | list Excel files without VBA signature | Workbook.HasMacro usage | VbaProject.IsSigned property | batch VBA security audit
// Common Searches: how to detect unsigned VBA macros in Excel files with Aspose.Cells | C# code to scan multiple workbooks for unsigned macro projects | Aspose.Cells find macro‑enabled workbooks lacking a digital signature | list Excel files with unsigned VBA using .NET
// Developer Intent: Locate Excel workbooks that contain macro projects without a digital signature and output their file names for further review or processing.
// Use Cases: Perform a security audit of macro‑enabled workbooks before distribution. | Generate compliance reports identifying files that need VBA signing. | Exclude unsigned macro workbooks from bulk data‑extraction pipelines.
// AI Prompts: Create a method that returns an array of paths for workbooks with unsigned VBA projects using Aspose.Cells. | Rewrite the example to write unsigned workbook names to a log file instead of the console. | Add support for password‑protected workbooks so the loop skips them without throwing an exception.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Iterates over a collection of workbook paths, loads each file with Aspose.Cells, checks Workbook.HasMacro and VbaProject.IsSigned, and writes the names of workbooks with unsigned VBA projects to the console while handling missing files and runtime errors.
    public class UnsignedVbaChecker
    {
        // Checks a collection of workbook files for unsigned VBA projects and logs their names.
        public static void Run(IEnumerable<string> workbookFiles)
        {
            foreach (var filePath in workbookFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file path.
                    Workbook workbook = new Workbook(filePath);

                    // Ensure the workbook actually contains a VBA project (macro-enabled).
                    if (workbook.HasMacro && workbook.VbaProject != null)
                    {
                        // If the VBA project is not signed, output the file name.
                        if (!workbook.VbaProject.IsSigned)
                        {
                            Console.WriteLine($"Unsigned VBA project detected: {filePath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }
    }

    public class Program
    {
        // Entry point of the application.
        public static void Main(string[] args)
        {
            try
            {
                var inputFiles = new List<string>();

                if (args.Length > 0)
                {
                    // Use command‑line arguments as file paths.
                    inputFiles.AddRange(args);
                }
                else
                {
                    // Prompt the user for file paths if none are provided.
                    Console.WriteLine("Enter workbook file paths separated by commas:");
                    var line = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        inputFiles.AddRange(line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries));
                    }
                }

                // Trim and filter existing files.
                var existingFiles = new List<string>();
                foreach (var path in inputFiles)
                {
                    var trimmed = path.Trim();
                    if (File.Exists(trimmed))
                    {
                        existingFiles.Add(trimmed);
                    }
                    else
                    {
                        Console.WriteLine($"File not found: {trimmed}");
                    }
                }

                // Run the unsigned VBA checker.
                UnsignedVbaChecker.Run(existingFiles);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
