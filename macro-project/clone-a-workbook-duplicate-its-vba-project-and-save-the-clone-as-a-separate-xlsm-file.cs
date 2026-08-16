// Title: Clone a macro‑enabled XLSM workbook and copy its VBA project with Aspose.Cells for .NET
// Description: Loads a source XLSM file, creates an empty workbook, copies all worksheets and VBA macros using CopyOptions.KeepMacros, optionally duplicates the VbaProject, and saves the clone as a new macro‑enabled XLSM file.
// Keywords: Aspose.Cells | .NET | C# | clone workbook | macro enabled | XLSM | VBA project | CopyOptions | KeepMacros | duplicate macros | Excel automation
// Common Searches: Aspose.Cells clone XLSM workbook with macros | Copy VBA project when duplicating Excel file in C# | Save cloned workbook as .xlsm using Aspose.Cells | KeepMacros option example Aspose.Cells .NET | How to duplicate a macro‑enabled workbook programmatically
// Developer Intent: Create an exact copy of a macro‑enabled workbook, preserving its VBA code, and store it as a separate XLSM file.
// Use Cases: Backing up a macro‑enabled workbook before automated processing. | Generating a reusable template that contains the same macros for multiple users. | Migrating VBA macros from one workbook to another within a data‑pipeline.
// AI Prompts: Generate C# code that clones an XLSM workbook and retains all VBA macros using Aspose.Cells. | Explain how CopyOptions.KeepMacros works and how to duplicate a VbaProject with Aspose.Cells. | Provide error‑handling best practices for missing source files and macro copy failures when cloning a macro‑enabled workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsCloneVba
{
    // Loads a source XLSM file, creates an empty workbook, copies all worksheets and VBA macros using CopyOptions.KeepMacros, optionally duplicates the VbaProject, and saves the clone as a new macro‑enabled XLSM file.
    class Program
    {
        static void Main()
        {
            // Path to the source macro‑enabled workbook
            string sourcePath = "source.xlsm";

            // Path where the cloned workbook will be saved
            string clonePath = "clone.xlsm";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the source workbook (contains VBA macros)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty workbook that will receive the clone
                Workbook clonedWorkbook = new Workbook();

                // Configure copy options to keep macros/VBA project
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy the entire content of the source workbook, including macros
                clonedWorkbook.Copy(sourceWorkbook, copyOptions);

                // Explicitly duplicate the VBA project (optional, reinforces the copy)
                if (sourceWorkbook.VbaProject != null && clonedWorkbook.VbaProject != null)
                {
                    clonedWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);
                }

                // Save the cloned workbook as a macro‑enabled file
                clonedWorkbook.Save(clonePath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook cloned successfully to '{clonePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
