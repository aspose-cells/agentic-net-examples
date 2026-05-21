using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVbaCheck
{
    public class UnsignedVbaChecker
    {
        // Checks a collection of workbook file paths for unsigned VBA projects.
        // Logs the file names of workbooks that contain macros but are not signed.
        public static void Run(IEnumerable<string> workbookPaths, string logFilePath)
        {
            try
            {
                // Ensure the log file is empty before starting.
                File.WriteAllText(logFilePath, string.Empty);

                foreach (string path in workbookPaths)
                {
                    // Skip non‑existent files.
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook.
                        Workbook workbook = new Workbook(path);

                        // Determine if the workbook contains a VBA project.
                        if (workbook.HasMacro)
                        {
                            // Check the signature status.
                            bool isSigned = workbook.VbaProject.IsSigned;

                            if (!isSigned)
                            {
                                // Log the unsigned workbook name.
                                string fileName = Path.GetFileName(path);
                                Console.WriteLine($"Unsigned VBA project: {fileName}");
                                File.AppendAllText(logFilePath, fileName + Environment.NewLine);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{path}': {ex.Message}");
                    }
                }

                Console.WriteLine($"Logging complete. Results saved to: {logFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point for the console application.
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: AsposeCellsVbaCheck <logFilePath> <workbookPath1> [<workbookPath2> ...]");
                    return;
                }

                string logFilePath = args[0];
                var workbookPaths = new List<string>(args.Length - 1);
                for (int i = 1; i < args.Length; i++)
                {
                    workbookPaths.Add(args[i]);
                }

                UnsignedVbaChecker.Run(workbookPaths, logFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}