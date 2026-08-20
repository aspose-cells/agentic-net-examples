// Title: Copy VBA chart‑generation macro from a template .xlsm to multiple workbooks using Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled template workbook, iterates through a list of target .xlsm files, copies the entire VBA project with VbaProject.Copy, saves each target as Xlsm to retain the macro, and disposes all workbook objects.
// Keywords: Aspose.Cells | C# | VBA project copy | macro‑enabled workbook | Xlsm | VbaProject.Copy | chart macro automation | batch macro copy
// Common Searches: Aspose.Cells copy VBA macro between .xlsm files | C# loop to copy chart macro to multiple workbooks | How to use VbaProject.Copy in Aspose.Cells | Batch copy VBA project to Excel files .NET | Copy macro from template workbook to many files
// Developer Intent: Duplicate a chart‑creating VBA macro from a template workbook to several macro‑enabled workbooks in a single loop.
// Use Cases: Distribute a standard chart‑generation macro across monthly report workbooks. | Create new reports on the fly and embed the same macro so every file has identical chart logic. | Update existing workbooks in a folder with the latest VBA macro version in one run.
// AI Prompts: Write C# code that uses Aspose.Cells to copy a VBA project from a template .xlsm to all .xlsm files in a given directory, with error handling and logging. | Show how to modify the example to copy only selected VBA modules instead of the whole project. | Provide an implementation that records each successful macro copy and captures detailed exceptions when VbaProject.Copy fails.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopyExample
{
    // Loads a macro‑enabled template workbook, iterates through a list of target .xlsm files, copies the entire VBA project with VbaProject.Copy, saves each target as Xlsm to retain the macro, and disposes all workbook objects.
    public class MacroCopier
    {
        public static void Run()
        {
            // Path to the macro‑enabled template workbook that contains the chart‑generating macro
            string templatePath = "TemplateWithMacro.xlsm";

            // Verify template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            Workbook templateWorkbook = null;
            try
            {
                // Load the template workbook (must be macro‑enabled)
                templateWorkbook = new Workbook(templatePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load template workbook: {ex.Message}");
                return;
            }

            // List of target workbook file paths (macro‑enabled extensions)
            List<string> targetPaths = new List<string>
            {
                "Target1.xlsm",
                "Target2.xlsm",
                "Target3.xlsm"
            };

            // Loop through each target workbook
            foreach (string targetPath in targetPaths)
            {
                Workbook targetWorkbook = null;
                try
                {
                    // Load existing target workbook or create a new one
                    if (File.Exists(targetPath))
                    {
                        targetWorkbook = new Workbook(targetPath);
                    }
                    else
                    {
                        targetWorkbook = new Workbook();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading/creating target workbook '{targetPath}': {ex.Message}");
                    continue;
                }

                try
                {
                    // Copy the VBA project (macro) from the template to the target workbook
                    targetWorkbook.VbaProject.Copy(templateWorkbook.VbaProject);

                    // Save the target workbook as a macro‑enabled file to preserve the copied macro
                    targetWorkbook.Save(targetPath, SaveFormat.Xlsm);
                    Console.WriteLine($"Macro copied and saved to {targetPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing target '{targetPath}': {ex.Message}");
                }
                finally
                {
                    targetWorkbook?.Dispose();
                }
            }

            // Clean up
            templateWorkbook?.Dispose();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                MacroCopier.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
