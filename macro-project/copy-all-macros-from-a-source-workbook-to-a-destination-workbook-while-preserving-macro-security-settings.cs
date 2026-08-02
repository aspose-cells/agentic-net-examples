// Title: Copy macros between .xlsm workbooks with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled source workbook, creates a new workbook, copies all worksheets and VBA code using CopyOptions.KeepMacros, transfers the source's macro security setting via Settings.EnableMacros, and saves the result as an Xlsm file.
// Keywords: Aspose.Cells | C# macro copy | CopyOptions KeepMacros | macro security Aspose | save Xlsm workbook | duplicate workbook with VBA
// Common Searches: Aspose.Cells copy macros from one .xlsm to another | preserve macro security when duplicating workbooks .NET | how to keep VBA code using Aspose.Cells CopyOptions | save copied workbook as macro‑enabled file
// Developer Intent: Programmatically duplicate a macro‑enabled workbook while retaining its VBA modules and original macro security configuration.
// Use Cases: Generate report files from a template that contains macros without losing functionality. | Create a secure backup of a macro‑enabled workbook for archival purposes. | Distribute a copy of a macro‑enabled spreadsheet to multiple users while preserving the original security level.
// AI Prompts: Write C# code with Aspose.Cells to copy all macros from a source .xlsm to a new workbook and keep the macro security flag. | Show error‑handling patterns for workbook copying when macros are involved using Aspose.Cells. | Explain how CopyOptions.KeepMacros and Settings.EnableMacros work together to maintain VBA code during workbook duplication.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroCopyDemo
{
    // Loads a macro‑enabled source workbook, creates a new workbook, copies all worksheets and VBA code using CopyOptions.KeepMacros, transfers the source's macro security setting via Settings.EnableMacros, and saves the result as an Xlsm file.
    public class Program
    {
        public static void Main()
        {
            const string sourcePath = "source_with_macros.xlsm";
            const string destPath = "destination_with_copied_macros.xlsm";

            try
            {
                // Ensure the source file exists before attempting to load it
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                    return;
                }

                // Load the source workbook that contains macros
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Preserve the macro security setting from the source workbook
                destinationWorkbook.Settings.EnableMacros = sourceWorkbook.Settings.EnableMacros;

                // Configure copy options to keep macros during the copy operation
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy the entire source workbook into the destination workbook,
                // including all worksheets, data, and macros
                sourceWorkbook.Copy(destinationWorkbook, copyOptions);

                // Save the result as a macro‑enabled workbook
                destinationWorkbook.Save(destPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(destPath)}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
