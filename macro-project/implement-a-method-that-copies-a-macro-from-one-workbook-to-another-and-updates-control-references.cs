// Title: Copy VBA macro between workbooks and sync references using Aspose.Cells for .NET
// Description: A C# helper that validates source and target files, loads both workbooks with Aspose.Cells, enables macros on the destination, copies the entire workbook while preserving VBA code (CopyOptions.KeepMacros), replicates VBA project references, and saves the result as a macro‑enabled .xlsm file.
// Keywords: Aspose.Cells copy macro | C# copy VBA workbook | KeepMacros option | sync VBA references | save as Xlsm | macro‑enabled workbook .NET | VBA project copy Aspose
// Common Searches: how to copy a macro from .xlsm to .xlsx using Aspose.Cells | Aspose.Cells copy VBA project references between workbooks | enable macros in destination workbook before saving as xlsm | transfer VBA modules with Aspose.Cells CopyOptions | C# copy macro and update references Aspose.Cells
// Developer Intent: Duplicate a macro‑enabled workbook into another file and update its VBA project references while keeping all macros intact.
// Use Cases: Apply a standard macro template to a newly generated report workbook. | Consolidate several macro‑enabled files into a single distribution workbook. | Refresh external library references after moving VBA code to a different workbook.
// AI Prompts: Generate C# code that copies selected VBA modules from one workbook to another with Aspose.Cells and adjusts reference paths. | Show robust error‑handling for macro copying, covering missing VBA projects, file‑access errors, and invalid paths. | Explain how to preserve custom forms, controls, and reference integrity when copying macros with Aspose.Cells, and how to verify the transfer.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroDemo
{
    // A C# helper that validates source and target files, loads both workbooks with Aspose.Cells, enables macros on the destination, copies the entire workbook while preserving VBA code (CopyOptions.KeepMacros), replicates VBA project references, and saves the result as a macro‑enabled .xlsm file.
    public static class MacroHelper
    {
        /// <param name="sourcePath">Path to the macro‑enabled source workbook (e.g., .xlsm).</param>
        /// <param name="destPath">Path to the destination workbook (can be a regular .xlsx file).</param>
        /// <param name="outputPath">Path where the resulting workbook will be saved (should be .xlsm to retain macros).</param>
        public static void CopyMacroAndUpdateReferences(string sourcePath, string destPath, string outputPath)
        {
            // Validate input files.
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");
            if (!File.Exists(destPath))
                throw new FileNotFoundException($"Destination workbook not found: {destPath}");

            try
            {
                // Load the source workbook that contains macros.
                using (Workbook sourceWorkbook = new Workbook(sourcePath))
                {
                    // Load the destination workbook (may or may not contain macros).
                    using (Workbook destWorkbook = new Workbook(destPath))
                    {
                        // Ensure the destination workbook is allowed to contain macros.
                        destWorkbook.Settings.EnableMacros = true;

                        // Configure copy options to keep macros during the copy operation.
                        CopyOptions copyOptions = new CopyOptions
                        {
                            KeepMacros = true
                        };

                        // Copy the entire source workbook into the destination workbook, preserving macros.
                        destWorkbook.Copy(sourceWorkbook, copyOptions);

                        // Synchronize VBA project references from source to destination.
                        if (destWorkbook.VbaProject != null && sourceWorkbook.VbaProject != null)
                        {
                            destWorkbook.VbaProject.References.Copy(sourceWorkbook.VbaProject.References);
                        }

                        // Save the resulting workbook as a macro‑enabled file.
                        destWorkbook.Save(outputPath, SaveFormat.Xlsm);
                    }
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to provide context.
                throw new InvalidOperationException("Failed to copy macros and update references.", ex);
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            // Example usage:
            string sourcePath = "SourceWithMacro.xlsm";
            string destPath = "EmptyWorkbook.xlsx";
            string outputPath = "ResultWithMacro.xlsm";

            try
            {
                MacroHelper.CopyMacroAndUpdateReferences(sourcePath, destPath, outputPath);
                Console.WriteLine($"Macro copy completed successfully. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
