// Title: Detect Unsigned VBA Projects in Excel Workbooks with Aspose.Cells for .NET
// Description: A C# utility that recursively scans a folder, loads each Excel file with Aspose.Cells, checks for VBA macros, determines if the VBA project is digitally signed, and outputs the paths of workbooks with unsigned macro projects.
// Keywords: Aspose.Cells unsigned VBA detection | C# scan Excel macros | list unsigned macro workbooks | VBA project signature check | recursive Excel file audit
// Common Searches: how to find unsigned VBA macros in Excel using Aspose.Cells | C# code to list Excel files with unsigned macro projects | scan directory for unsigned VBA projects .xlsm | detect unsigned VBA projects in workbooks Aspose.Cells
// Developer Intent: Locate every Excel workbook in a specified directory that contains a VBA project lacking a digital signature.
// Use Cases: Security audit of shared drives to identify potentially unsafe macro‑enabled files. | Pre‑migration report of workbooks that may require macro signing or removal. | Automated gate in CI/CD pipelines that blocks builds when unsigned VBA macros are present.
// AI Prompts: Generate a method returning List<string> of file paths for unsigned VBA projects using Aspose.Cells. | Extend the example to export the list of unsigned files to a CSV or JSON report. | Add structured logging (e.g., Serilog) that records each processed workbook and its signature status.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# utility that recursively scans a folder, loads each Excel file with Aspose.Cells, checks for VBA macros, determines if the VBA project is digitally signed, and outputs the paths of workbooks with unsigned macro projects.
    public class DetectUnsignedVbaProjects
    {
        public static void Run(string directoryPath)
        {
            // Verify that the directory exists
            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory does not exist: {directoryPath}");
                return;
            }

            // Collection to store file names that contain unsigned VBA projects
            List<string> unsignedFiles = new List<string>();

            // Excel file extensions that Aspose.Cells can handle
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xls2003", ".xls2007" };

            // Scan all files recursively in the specified directory
            foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.*", SearchOption.AllDirectories))
            {
                // Skip files that are not Excel workbooks
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue;

                // Ensure the file actually exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Proceed only if the workbook actually contains VBA macros
                    if (workbook.HasMacro)
                    {
                        // Determine whether the VBA project is signed
                        bool isSigned = workbook.VbaProject.IsSigned;

                        // If the VBA project is not signed, record the file name
                        if (!isSigned)
                        {
                            unsignedFiles.Add(filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors (e.g., corrupted file) and continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Output the list of files with unsigned VBA projects
            Console.WriteLine("Files with unsigned VBA projects:");
            foreach (string file in unsignedFiles)
            {
                Console.WriteLine(file);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            string directoryPath;

            if (args.Length > 0)
            {
                directoryPath = args[0];
            }
            else
            {
                Console.Write("Enter directory path to scan: ");
                directoryPath = Console.ReadLine();
            }

            try
            {
                DetectUnsignedVbaProjects.Run(directoryPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
