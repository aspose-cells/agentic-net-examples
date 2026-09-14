// Title: Batch export VBA modules from .xlsm workbooks to individual .bas files using Aspose.Cells for .NET
// AI Prompts: Write a C# method that accepts a collection of .xlsm file paths and an output folder, loads each workbook with Aspose.Cells, and saves every VBA module as a separate .bas file named "<Workbook>_<Module>.bas". | Enhance the export routine to prepend a comment header to each .bas file that records the source workbook name and module name, and ensure any illegal filename characters are replaced. | Add comprehensive error handling and logging to the VBA extraction process: record missing workbook files, workbooks without a VBA project, and any file‑write exceptions in a log file.
// Common Searches: how to programmatically extract VBA code from .xlsm files using Aspose.Cells in C# | C# batch export of VBA modules to .bas files from multiple Excel workbooks | save each VBA module from an Excel macro‑enabled workbook as a separate file with Aspose.Cells | automate VBA project extraction for version control using Aspose.Cells .NET | Aspose.Cells example for exporting VBA modules to text files
// Tags: extract VBA code to .bas files using Aspose.Cells | batch process .xlsm workbooks for VBA module extraction C# | sanitize VBA module filenames for file system compatibility | log missing VBA projects during workbook processing Aspose.Cells | manage VBA project version control with Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace VbaExportExample
{
    // The sample program iterates over a list of .xlsm workbook paths, creates the target directory if needed, loads each workbook with Aspose.Cells, checks for a VBA project, and writes the source code of each VBA module to a uniquely named .bas file (WorkbookName_ModuleName.bas). It handles missing files, load failures, absent VBA projects, and write errors, reporting status via console output.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide a list of workbook file paths and an output directory.
            List<string> workbookPaths = new List<string>
            {
                @"C:\Workbooks\Sample1.xlsm",
                @"C:\Workbooks\Sample2.xlsm"
                // Add more workbook paths as needed.
            };

            string outputDirectory = @"C:\VbaExports";

            try
            {
                ExportVbaModules(workbookPaths, outputDirectory);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <param name="workbookPaths">List of full paths to the workbooks.</param>
        /// <param name="outputDir">Directory where .bas files will be saved.</param>
        static void ExportVbaModules(IEnumerable<string> workbookPaths, string outputDir)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            foreach (string wbPath in workbookPaths)
            {
                // Verify the workbook file exists before attempting to load.
                if (!File.Exists(wbPath))
                {
                    Console.WriteLine($"Workbook file not found: {wbPath}");
                    continue;
                }

                Workbook workbook;
                try
                {
                    // Load the workbook.
                    workbook = new Workbook(wbPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook '{wbPath}': {ex.Message}");
                    continue;
                }

                // Check if the workbook contains a VBA project.
                if (workbook.VbaProject == null || workbook.VbaProject.Modules == null)
                {
                    Console.WriteLine($"No VBA project found in workbook: {wbPath}");
                    continue;
                }

                // Iterate through each VBA module.
                foreach (var module in workbook.VbaProject.Modules)
                {
                    // Build a safe file name: WorkbookName_ModuleName.bas
                    string workbookName = Path.GetFileNameWithoutExtension(wbPath);
                    string moduleName = module.Name;

                    // Replace any invalid filename characters.
                    foreach (char invalidChar in Path.GetInvalidFileNameChars())
                    {
                        moduleName = moduleName.Replace(invalidChar, '_');
                    }

                    string basFileName = $"{workbookName}_{moduleName}.bas";
                    string basFilePath = Path.Combine(outputDir, basFileName);

                    try
                    {
                        // Write the module's code to the .bas file.
                        File.WriteAllText(basFilePath, module.Codes ?? string.Empty);
                        Console.WriteLine($"Exported module '{moduleName}' from '{workbookName}' to '{basFilePath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write module '{moduleName}' to file: {ex.Message}");
                    }
                }
            }
        }
    }
}
