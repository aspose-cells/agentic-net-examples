// Title: Copy a VBA macro from a macro‑enabled template workbook to multiple Excel files using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsm template, copies its VBA project into each workbook from a list of .xlsx files, and saves the results as .xlsm files with Aspose.Cells. | Write a loop that merges the worksheets of target workbooks into a fresh copy of the macro‑enabled template while preserving the VBA project, and handle missing files gracefully. | Add comprehensive error handling and logging to the macro‑copy process to ensure the VBA code remains intact after each workbook is saved.
// Common Searches: how to copy VBA macro from a template workbook to several Excel files using Aspose.Cells C# | batch add a macro to multiple .xlsx files and save as .xlsm with Aspose.Cells | preserve VBA project when merging worksheets into a macro‑enabled workbook in .NET | Aspose.Cells example for cloning a workbook and keeping its VBA code | C# loop to process a list of Excel files and attach a macro template
// Tags: copy VBA project Aspose.Cells | macro‑enabled workbook generation .NET | batch process Excel files Aspose.Cells | merge worksheets into Xlsm using Aspose | preserve VBA when cloning workbook | load macro template Aspose.Cells | save workbook as Xlsm Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads a macro‑enabled .xlsm template, verifies it contains a VBA project, then iterates over a list of target .xlsx files. For each target it creates a fresh copy of the template, clears its original sheets, copies the target's worksheets into the copy, and saves the result as a new .xlsm file, preserving the original macro.
class Program
{
    static void Main()
    {
        // Path to the macro‑enabled template workbook
        string templatePath = "TemplateWithMacro.xlsm";

        // Ensure the template file exists before loading
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        try
        {
            // Load the template workbook (must be .xlsm)
            Workbook templateWorkbook = new Workbook(templatePath);

            // Verify that the template contains a VBA project
            if (templateWorkbook.VbaProject == null)
            {
                Console.WriteLine("The template workbook does not contain a VBA project.");
                return;
            }

            // List of target workbooks to which the macro will be added
            List<string> targetWorkbookPaths = new List<string>
            {
                "Target1.xlsx",
                "Target2.xlsx",
                "Target3.xlsx"
            };

            foreach (string targetPath in targetWorkbookPaths)
            {
                // Verify target file existence
                if (!File.Exists(targetPath))
                {
                    Console.WriteLine($"Target file not found: {targetPath}");
                    continue;
                }

                try
                {
                    // Load the target workbook (may be .xlsx without macros)
                    Workbook targetWorkbook = new Workbook(targetPath);

                    // Load a fresh copy of the template for each target to retain the macro
                    Workbook resultWorkbook = new Workbook(templatePath);

                    // Remove all worksheets from the copied template
                    resultWorkbook.Worksheets.Clear();

                    // Copy worksheets from the target workbook into the result workbook
                    foreach (Worksheet ws in targetWorkbook.Worksheets)
                    {
                        // AddCopy expects the source worksheet name
                        resultWorkbook.Worksheets.AddCopy(ws.Name);
                    }

                    // Save as macro‑enabled workbook (.xlsm) to retain the macro
                    string outputPath = Path.ChangeExtension(targetPath, ".xlsm");
                    resultWorkbook.Save(outputPath, SaveFormat.Xlsm);
                    Console.WriteLine($"Saved macro‑enabled workbook: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{targetPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load template workbook: {ex.Message}");
        }
    }
}
