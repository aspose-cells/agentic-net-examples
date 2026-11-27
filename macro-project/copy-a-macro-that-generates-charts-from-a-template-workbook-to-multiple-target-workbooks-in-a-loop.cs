// Title: Copy a VBA macro from a template workbook to multiple Excel files using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy VBA macro from a template to several workbooks | how to add a macro-enabled VBA project to multiple Excel files in a loop with Aspose.Cells | batch copy chart‑generating macro to .xlsx files and save as .xlsm using C# | preserve VBA project while copying worksheets with Aspose.Cells .NET | automate creation of macro‑enabled workbooks for many files using Aspose.Cells
// Tags: copy VBA macro Aspose.Cells C# | add VBA project to existing workbook Aspose.Cells | generate Xlsm from Xlsx Aspose.Cells | batch process workbooks with VBA project Aspose.Cells | preserve VBA when copying worksheets Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a macro‑enabled .xlsm template, iterates over a list of target .xlsx workbooks, creates a new workbook that inherits the VBA project from the template, clears default sheets, copies each worksheet from the target workbook, and saves the combined result as a macro‑enabled .xlsm file.
class MacroCopyExample
{
    static void Main()
    {
        // Path to the template workbook that contains the macro (must be a macro‑enabled file)
        string templatePath = "TemplateWithMacro.xlsm";

        // Verify that the template file exists
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // List of target workbook file paths (can be .xlsx or any other supported format)
        List<string> targetPaths = new List<string>
        {
            "Target1.xlsx",
            "Target2.xlsx",
            "Target3.xlsx"
        };

        // Load the template workbook once (it contains the VBA project)
        Workbook templateWb;
        try
        {
            templateWb = new Workbook(templatePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load template workbook: {ex.Message}");
            return;
        }

        // Process each target workbook
        foreach (string targetPath in targetPaths)
        {
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target file not found: {targetPath}");
                continue;
            }

            try
            {
                // Load the target workbook (no macro needed at load time)
                Workbook targetWb = new Workbook(targetPath);

                // Create a new workbook and copy the template (preserves the VBA project)
                Workbook resultWb = new Workbook();
                templateWb.Copy(resultWb);

                // Remove existing worksheets from the copy
                resultWb.Worksheets.Clear();

                // Copy worksheets from the target workbook into the result workbook
                foreach (Worksheet ws in targetWb.Worksheets)
                {
                    // Add a copy of each worksheet by name
                    resultWb.Worksheets.AddCopy(ws.Name);
                }

                // Determine the output file name (save as macro‑enabled .xlsm)
                string outputPath = Path.ChangeExtension(targetPath, ".xlsm");

                // Save the result workbook with the copied macro
                resultWb.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Successfully created: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{targetPath}': {ex.Message}");
            }
        }
    }
}
