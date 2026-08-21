// Title: C# batch scan for unsigned VBA projects in Excel workbooks with Aspose.Cells
// Description: A C# utility that iterates over a list of Excel files, loads each workbook with Aspose.Cells, checks for macro presence, evaluates the VbaProject.IsSigned flag, and logs the paths of workbooks that contain unsigned VBA projects. The program also reports files without macros, handles missing files, and captures processing errors.
// Keywords: Aspose.Cells unsigned VBA detection | C# batch macro signing check | Excel workbook VbaProject.IsSigned | detect unsigned macros .NET | security audit Excel macros | macro signing compliance | CI pipeline VBA validation | list Excel files without signed VBA | Aspose.Cells macro security
// Common Searches: how to find unsigned VBA projects in Excel using Aspose.Cells | C# code to list workbooks with unsigned macros | batch check macro signing status in .NET | Aspose.Cells detect unsigned VBA projects | scan multiple Excel files for unsigned macros
// Developer Intent: Locate Excel workbooks that contain VBA projects lacking a digital signature and output their file names.
// Use Cases: Security audit to flag macro‑enabled files before distribution | Compliance reporting of unsigned VBA projects across a document repository | Automated validation step in CI/CD pipelines to prevent unsigned macros from being released
// AI Prompts: Create a method that returns a List<string> of paths with unsigned VBA projects instead of writing to the console. | Modify the program to export unsigned workbook details to a CSV file with columns for file path, signing status, and timestamp. | Add recursive directory traversal to discover all Excel files and apply the unsigned VBA check automatically.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    // A C# utility that iterates over a list of Excel files, loads each workbook with Aspose.Cells, checks for macro presence, evaluates the VbaProject.IsSigned flag, and logs the paths of workbooks that contain unsigned VBA projects. The program also reports files without macros, handles missing files, and captures processing errors.
    public class UnsignedVbaChecker
    {
        /// <param name="workbookPaths">Collection of full file paths to Excel workbooks.</param>
        public static void Run(IEnumerable<string> workbookPaths)
        {
            foreach (string path in workbookPaths)
            {
                try
                {
                    // Verify that the file exists before attempting to load
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        continue;
                    }

                    // Load the workbook from the file system
                    Workbook workbook = new Workbook(path);

                    // Determine whether the workbook contains any VBA/macros
                    if (workbook.HasMacro)
                    {
                        // Access the VBA project associated with the workbook
                        VbaProject vbaProject = workbook.VbaProject;

                        // Check the signing status; IsSigned is true when the project is signed
                        bool isSigned = vbaProject.IsSigned;

                        if (!isSigned)
                        {
                            // Log the file name of the workbook with an unsigned VBA project
                            Console.WriteLine($"Unsigned VBA project detected: {path}");
                        }
                    }
                    else
                    {
                        // Optional: log workbooks that do not contain any VBA at all
                        Console.WriteLine($"No VBA macro present: {path}");
                    }
                }
                catch (Exception ex)
                {
                    // Log any unexpected errors for the current file
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                // If no arguments are provided, display usage information
                if (args == null || args.Length == 0)
                {
                    Console.WriteLine("Usage: AsposeCellsVbaCheck <full_path_to_excel_file1> [<full_path_to_excel_file2> ...]");
                    return;
                }

                // Run the unsigned VBA checker on the supplied file paths
                UnsignedVbaChecker.Run(args);
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors at the top level
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
