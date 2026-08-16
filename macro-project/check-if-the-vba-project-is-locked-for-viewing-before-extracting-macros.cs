// Title: C# – Detect VBA Project Protection and Extract Macros with Aspose.Cells for .NET
// Description: Loads an .xlsm workbook, confirms the presence of macros, checks the VBA project's IsProtected flag, and, when unprotected, iterates through each VbaModule to output its name, type, and source code. Includes handling for missing files, workbooks without macros, and protected projects.
// Keywords: Aspose.Cells VBA protection | C# extract macros from xlsm | VbaProject.IsProtected | read VBA modules Aspose.Cells | check workbook.HasMacro | .NET Excel macro extraction
// Common Searches: how to know if a VBA project is locked using Aspose.Cells | C# code to list VBA modules only when not protected | Aspose.Cells check workbook.HasMacro and VbaProject.IsProtected | extract macro source from .xlsm with Aspose.Cells .NET
// Developer Intent: Identify whether an Excel workbook's VBA project is locked for viewing and, if it is not, retrieve the macro code from each module.
// Use Cases: Pre‑flight validation of macro accessibility before bulk analysis of workbooks. | Skipping password‑protected VBA projects during automated macro extraction pipelines. | Logging unprotected VBA source for compliance audits or documentation.
// AI Prompts: Create a C# function using Aspose.Cells that returns true when a workbook's VBA project is protected. | Generate code that saves each VBA module from an unprotected workbook as a separate .vba file. | Provide comprehensive error handling for missing files, workbooks without macros, and protected VBA projects when extracting macros with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads an .xlsm workbook, confirms the presence of macros, checks the VBA project's IsProtected flag, and, when unprotected, iterates through each VbaModule to output its name, type, and source code. Includes handling for missing files, workbooks without macros, and protected projects.
    public class CheckVbaLockAndExtractMacros
    {
        public static void Run()
        {
            const string filePath = "sample.xlsm";

            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Verify the workbook actually contains VBA/macros
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any macros.");
                    return;
                }

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("Unable to access the VBA project.");
                    return;
                }

                // Check if the VBA project is protected (locked) for viewing
                if (vbaProject.IsProtected)
                {
                    Console.WriteLine("The VBA project is protected. Cannot extract macros.");
                    return;
                }

                // If not protected, iterate through all VBA modules and output their code
                Console.WriteLine("Extracting macros from the VBA project:");
                for (int i = 0; i < vbaProject.Modules.Count; i++)
                {
                    VbaModule module = vbaProject.Modules[i];
                    Console.WriteLine($"--- Module: {module.Name} (Type: {module.Type}) ---");
                    Console.WriteLine(module.Codes);
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaLockAndExtractMacros.Run();
        }
    }
}
