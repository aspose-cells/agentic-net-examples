// Title: Check VBA Project Protection and Extract Macros from an XLSM using Aspose.Cells for .NET
// Description: Loads a macro‑enabled workbook, verifies the presence of a VBA project, tests the VbaProject.IsProtected flag, and when unprotected iterates through each VbaModule to display and save its code as a .bas file while handling missing files and I/O exceptions.
// Keywords: Aspose.Cells VBA protection | VbaProject.IsProtected .NET | extract VBA modules C# | save VBA code to .bas | macro‑enabled workbook processing | locked VBA project detection
// Common Searches: how to detect locked VBA project with Aspose.Cells | extract VBA code from xlsm using C# | skip macro extraction when VBA project is password protected | save each VBA module as .bas file Aspose.Cells | check if Excel VBA project is protected programmatically
// Developer Intent: Identify whether an XLSM file’s VBA project is password‑protected and, if it is not, export each macro module to a separate .bas file.
// Use Cases: Pre‑flight validation of workbooks in a batch job to avoid errors on protected VBA projects. | Automated migration of VBA code to source control by exporting modules from multiple files. | Integration into an Excel audit tool that logs macro content only for unprotected projects.
// AI Prompts: Generate C# code that uses Aspose.Cells to list VBA module names without extracting their source when the project is protected. | Create a method that extracts VBA modules to .bas files only if VbaProject.IsProtected is false, with robust file‑I/O error handling. | Refactor the demo to write extraction results to a log file and return a collection of the saved .bas file paths.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled workbook, verifies the presence of a VBA project, tests the VbaProject.IsProtected flag, and when unprotected iterates through each VbaModule to display and save its code as a .bas file while handling missing files and I/O exceptions.
    public class VbaProjectLockCheckAndExtractDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the macro‑enabled Excel file
                string inputPath = "sample.xlsm";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook contains a VBA project
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any VBA project.");
                    return;
                }

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("Unable to retrieve the VBA project.");
                    return;
                }

                // Check if the VBA project is protected (locked for viewing)
                if (vbaProject.IsProtected)
                {
                    Console.WriteLine("The VBA project is protected. Cannot extract macros.");
                    return;
                }

                // Extract macro code from each module
                Console.WriteLine("Extracting VBA modules...");

                for (int i = 0; i < vbaProject.Modules.Count; i++)
                {
                    VbaModule module = vbaProject.Modules[i];
                    string moduleName = module.Name;
                    string moduleCode = module.Codes;

                    Console.WriteLine($"--- Module: {moduleName} ---");
                    Console.WriteLine(moduleCode);
                    Console.WriteLine();

                    // Save each module's code to a .bas file
                    string outputFile = $"{moduleName}.bas";
                    try
                    {
                        File.WriteAllText(outputFile, moduleCode);
                        Console.WriteLine($"Module code saved to: {outputFile}");
                    }
                    catch (Exception writeEx)
                    {
                        Console.WriteLine($"Failed to write module file '{outputFile}': {writeEx.Message}");
                    }
                }

                Console.WriteLine("Macro extraction completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectLockCheckAndExtractDemo.Run();
        }
    }
}
