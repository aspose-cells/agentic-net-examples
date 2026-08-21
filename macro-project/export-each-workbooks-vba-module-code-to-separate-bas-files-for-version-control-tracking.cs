// Title: Export VBA Modules to .bas Files from .xlsm Workbooks with Aspose.Cells (C#)
// Description: C# utility that scans a directory for .xlsm files, loads each workbook with Aspose.Cells, detects VBA projects, creates a workbook‑named subfolder, and writes every VBA module to a sanitized .bas file for easy version‑control tracking.
// Keywords: Aspose.Cells VBA export | C# extract VBA modules | save VBA code .bas | batch export .xlsm macros | .NET VBA project extraction | version control Excel macros | export VBA to files | Aspose.Cells macro handling
// Common Searches: export VBA modules from xlsm using Aspose.Cells C# | save each VBA module as .bas file | batch extract Excel macros for source control | C# code to write VBA project to files | Aspose.Cells export macro code
// Developer Intent: Generate separate .bas files for every VBA module in each macro‑enabled workbook to enable source‑control of macro code.
// Use Cases: Automated nightly extraction of all VBA modules from a repository of .xlsm files. | Creating a per‑workbook folder structure that isolates macro code for Git tracking. | Integrating the export routine into CI/CD pipelines to capture macro changes alongside source code.
// AI Prompts: Write a C# program that uses Aspose.Cells to iterate over .xlsm files in a folder and export each VBA module to a .bas file, handling invalid filename characters. | Refactor ExportVbaModules to return a list of exported file paths and allow the caller to specify a custom file extension. | Create a PowerShell script that calls the compiled .NET assembly, logs successes and failures, and uploads the exported .bas files to a Git repository.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaExportExample
{
    // C# utility that scans a directory for .xlsm files, loads each workbook with Aspose.Cells, detects VBA projects, creates a workbook‑named subfolder, and writes every VBA module to a sanitized .bas file for easy version‑control tracking.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: export VBA modules from all .xlsm files in a folder
            string sourceFolder = @"C:\Workbooks";
            string outputFolder = @"C:\VbaExports";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each macro-enabled workbook in the source folder
            foreach (string workbookPath in Directory.GetFiles(sourceFolder, "*.xlsm"))
            {
                ExportVbaModules(workbookPath, outputFolder);
            }

            Console.WriteLine("Export completed.");
        }

        /// <param name="workbookPath">Full path to the macro-enabled workbook.</param>
        /// <param name="outputFolder">Folder where .bas files will be saved.</param>
        static void ExportVbaModules(string workbookPath, string outputFolder)
        {
            // Load the workbook (macro-enabled format)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project; if none exists, skip this workbook
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null || vbaProject.Modules.Count == 0)
            {
                Console.WriteLine($"No VBA project found in '{Path.GetFileName(workbookPath)}'.");
                return;
            }

            // Create a subfolder named after the workbook (without extension) to hold its modules
            string workbookName = Path.GetFileNameWithoutExtension(workbookPath);
            string workbookExportFolder = Path.Combine(outputFolder, workbookName);
            Directory.CreateDirectory(workbookExportFolder);

            // Iterate through all modules in the VBA project
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                VbaModule module = vbaProject.Modules[i];

                // Determine a safe file name for the module
                string moduleFileName = $"{module.Name}.bas";
                foreach (char invalid in Path.GetInvalidFileNameChars())
                {
                    moduleFileName = moduleFileName.Replace(invalid, '_');
                }

                string moduleFilePath = Path.Combine(workbookExportFolder, moduleFileName);

                // Write the module's code to the .bas file
                File.WriteAllText(moduleFilePath, module.Codes ?? string.Empty);
                Console.WriteLine($"Exported module '{module.Name}' to '{moduleFilePath}'.");
            }
        }
    }
}
