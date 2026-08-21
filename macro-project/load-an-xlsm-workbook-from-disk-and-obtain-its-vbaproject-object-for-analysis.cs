// Title: C# – Load an XLSM workbook and access its VbaProject with Aspose.Cells for .NET
// Description: Demonstrates how to open a macro‑enabled Excel file (XLSM) from disk using Aspose.Cells, verify the file's existence, retrieve the workbook's VbaProject, and display key VBA metadata such as project name, signing status, protection flag, and module count. Includes graceful handling for missing files or workbooks without VBA projects.
// Keywords: Aspose.Cells C# | load XLSM workbook | VbaProject extraction | macro‑enabled Excel | VBA project properties | Aspose.Cells VBA analysis | .NET Excel macro example | read VBA modules
// Common Searches: Aspose.Cells get VbaProject from XLSM C# | How to read VBA project details with Aspose.Cells | C# sample to list VBA modules in a macro‑enabled workbook | Check if Excel file contains signed VBA macros using Aspose.Cells | Retrieve VBA project name and module count with Aspose.Cells
// Developer Intent: Open a macro‑enabled Excel file and obtain its VbaProject to inspect VBA metadata.
// Use Cases: Display VBA project name, signing status, protection flag, and module count for an XLSM workbook. | Validate the presence of a VBA project before performing macro analysis or transformation. | Provide clear error messages when the target file is missing or lacks a VBA project.
// AI Prompts: Generate C# code with Aspose.Cells that lists all VBA module names in an XLSM file. | Create a method to extract and save the source code of each VBA module from a loaded workbook. | Suggest robust error‑handling patterns for accessing VbaProject when the file may be absent or contain no macros.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    // Demonstrates how to open a macro‑enabled Excel file (XLSM) from disk using Aspose.Cells, verify the file's existence, retrieve the workbook's VbaProject, and display key VBA metadata such as project name, signing status, protection flag, and module count. Includes graceful handling for missing files or workbooks without VBA projects.
    public class LoadVbaProject
    {
        public static void Run()
        {
            // Path to the macro-enabled workbook (XLSM) on disk
            string workbookPath = "sample_with_macro.xlsm";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook using the string constructor (provided rule)
                Workbook workbook = new Workbook(workbookPath);

                // Obtain the VbaProject object from the loaded workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Simple analysis: display some basic VBA project information
                if (vbaProject != null)
                {
                    Console.WriteLine("VBA Project Name: " + vbaProject.Name);
                    Console.WriteLine("Is Signed: " + vbaProject.IsSigned);
                    Console.WriteLine("Is Protected: " + vbaProject.IsProtected);
                    Console.WriteLine("Modules Count: " + vbaProject.Modules.Count);
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }

        // Entry point required for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
