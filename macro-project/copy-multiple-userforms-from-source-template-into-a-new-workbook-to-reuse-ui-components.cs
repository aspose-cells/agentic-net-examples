// Title: Copy VBA UserForms from a macro‑enabled template to a new workbook with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled .xlsm template containing VBA UserForms, ensures the destination folder exists, and saves a new .xlsm workbook while preserving the entire VBA project and all UserForms using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | copy VBA UserForms | macro enabled workbook | xlsm | preserve VBA project | Excel UI automation | programmatic VBA copy | Excel workbook cloning | VBA forms transfer
// Common Searches: Aspose.Cells copy UserForms .xlsm | How to duplicate VBA UserForms with C# | Preserve macros when saving workbook Aspose.Cells | Copy macro‑enabled Excel template programmatically | Transfer VBA forms between workbooks .NET
// Developer Intent: Copy the VBA UserForms and full VBA project from a source macro‑enabled workbook into a new workbook using Aspose.Cells for .NET.
// Use Cases: Generate multiple reports that share the same UI components defined in a template's UserForms. | Automate creation of standardized Excel files for different departments while retaining all macros and forms. | Deploy a common Excel front‑end across a suite of applications by cloning the template’s VBA project into each generated file.
// AI Prompts: Show C# code to copy selected UserForms from a source .xlsm to a destination workbook with Aspose.Cells. | Explain how to merge VBA projects from two workbooks while handling duplicate UserForm names using Aspose.Cells. | Provide guidance on preserving macro security settings when saving a workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUserFormCopyDemo
{
    // Loads a macro‑enabled .xlsm template containing VBA UserForms, ensures the destination folder exists, and saves a new .xlsm workbook while preserving the entire VBA project and all UserForms using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            const string sourcePath = "SourceTemplate.xlsm";
            const string destPath = "DestinationWithUserForms.xlsm";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the source workbook that contains the UserForms (must be a macro‑enabled file)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Ensure the destination directory exists
                string destDir = Path.GetDirectoryName(destPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the workbook to the destination path (preserves VBA project and UserForms)
                sourceWorkbook.Save(destPath, SaveFormat.Xlsm);

                Console.WriteLine("UserForms copied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
