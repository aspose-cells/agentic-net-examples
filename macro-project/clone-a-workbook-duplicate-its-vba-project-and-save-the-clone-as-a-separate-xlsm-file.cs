// Title: Clone a macro‑enabled XLSM workbook and copy its VBA project with Aspose.Cells for .NET (C#)
// Description: C# example that loads a source XLSM file, creates an empty workbook, uses CopyOptions.KeepMacros to copy all worksheets, duplicates the VbaProject via VbaProject.Copy, and saves the clone as a new macro‑enabled workbook. Includes file‑existence checks and exception handling.
// Keywords: Aspose.Cells clone XLSM | CopyOptions KeepMacros C# | VbaProject.Copy Aspose | macro enabled workbook duplicate | C# Aspose.Cells VBA cloning | save as Xlsm Aspose
// Common Searches: how to copy an XLSM workbook with macros using Aspose.Cells | Aspose.Cells duplicate VBA project C# | clone macro‑enabled Excel file .NET | copy worksheets and VBA code with Aspose.Cells | C# example for preserving macros when cloning Excel
// Developer Intent: Create an exact copy of a macro‑enabled Excel workbook, including its VBA project, and store the clone as a separate XLSM file.
// Use Cases: Backup a macro‑enabled workbook before applying automated changes. | Generate a reusable template that contains the same VBA macros for distribution. | Implement version control for Excel files with embedded VBA by cloning the original for each release.
// AI Prompts: Write C# code that clones an XLSM workbook, copies its VBA project, and saves the result as a new XLSM file using Aspose.Cells, with error handling for missing source files. | Explain how CopyOptions.KeepMacros and VbaProject.Copy work together to preserve VBA code when duplicating a workbook in Aspose.Cells. | Show how to modify the cloning process to skip selected worksheets while still copying the full VBA project.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsCloneDemo
{
    // C# example that loads a source XLSM file, creates an empty workbook, uses CopyOptions.KeepMacros to copy all worksheets, duplicates the VbaProject via VbaProject.Copy, and saves the clone as a new macro‑enabled workbook. Includes file‑existence checks and exception handling.
    class Program
    {
        static void Main()
        {
            // Paths for source and cloned workbooks
            string sourcePath = "source.xlsm";
            string clonePath = "clone.xlsm";

            try
            {
                // Verify that the source file exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                    return;
                }

                // Load the source macro‑enabled workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty workbook for the clone
                Workbook clonedWorkbook = new Workbook();

                // Set copy options to retain macros/VBA code
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy the entire content of the source workbook into the clone
                sourceWorkbook.Copy(clonedWorkbook, copyOptions);

                // Duplicate the VBA project from source to clone
                if (sourceWorkbook.VbaProject != null && clonedWorkbook.VbaProject != null)
                {
                    clonedWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);
                }

                // Save the cloned workbook as a macro‑enabled file
                clonedWorkbook.Save(clonePath, SaveFormat.Xlsm);
                Console.WriteLine($"Cloned workbook saved to: {Path.GetFullPath(clonePath)}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
