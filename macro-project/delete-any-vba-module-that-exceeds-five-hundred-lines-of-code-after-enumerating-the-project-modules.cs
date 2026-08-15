// Title: Delete VBA Modules Over 500 Lines in an .xlsm Workbook with Aspose.Cells for .NET
// Description: Loads a macro‑enabled Excel file, enumerates its VBA project, counts the lines in each module, removes modules whose code exceeds 500 lines, and saves the cleaned workbook.
// Keywords: Aspose.Cells | C# VBA module removal | delete large VBA module | enumerate VBA project | count VBA lines | macro-enabled workbook cleanup | Aspose.Cells .NET | remove VBA from xlsm
// Common Searches: Aspose.Cells delete VBA module over 500 lines | C# remove large VBA modules from .xlsm | how to count lines in VBA module using Aspose | enumerate VBA modules Aspose.Cells | programmatically clean VBA code in Excel workbook | remove VBA modules with Aspose.Cells .NET
// Developer Intent: Programmatically eliminate any VBA module that contains more than 500 lines from a macro‑enabled Excel workbook.
// Use Cases: Reduce workbook size and improve performance by stripping oversized macro modules before distribution. | Enforce compliance by automatically removing large VBA sections from generated reports in batch jobs. | Maintain a consistent macro footprint during automated workbook creation by deleting lengthy modules.
// AI Prompts: Generate C# code using Aspose.Cells that lists all VBA modules and deletes those with a line count greater than 500. | Show how to add a configurable line‑limit parameter and log the names of removed modules in the example. | Recommend extra error handling and resource cleanup for the VBA module removal workflow.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled Excel file, enumerates its VBA project, counts the lines in each module, removes modules whose code exceeds 500 lines, and saves the cleaned workbook.
    public class DeleteLargeVbaModules
    {
        public static void Run()
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.xlsm";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load an existing macro-enabled workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project and its module collection
                VbaProject vbaProject = workbook.VbaProject;
                VbaModuleCollection modules = vbaProject.Modules;

                // Collect names of modules whose code exceeds 500 lines
                List<string> modulesToRemove = new List<string>();

                for (int i = 0; i < modules.Count; i++)
                {
                    VbaModule module = modules[i];
                    string code = module.Codes ?? string.Empty;

                    // Count lines by splitting on newline characters
                    int lineCount = code.Split(new[] { '\n' }, StringSplitOptions.None).Length;

                    if (lineCount > 500)
                    {
                        modulesToRemove.Add(module.Name);
                    }
                }

                // Remove the identified modules by name
                foreach (string name in modulesToRemove)
                {
                    modules.Remove(name);
                }

                // Save the workbook after removal
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteLargeVbaModules.Run();
        }
    }
}
