// Title: Copy VBA macros between Excel workbooks with Aspose.Cells for .NET
// Description: Shows how to verify source and destination files, load them using Aspose.Cells, copy all content while preserving VBA via CopyOptions.KeepMacros, enable macros in the target workbook, and save the result as a macro‑enabled .xlsm file.
// Keywords: Aspose.Cells copy macros | C# copy VBA between workbooks | CopyOptions KeepMacros | macro‑enabled .xlsm save | preserve VBA Aspose | Excel macro transfer .NET | update form controls Aspose.Cells | merge workbooks with macros
// Common Searches: Aspose.Cells copy VBA from one workbook to another | How to keep macros when merging Excel files using C# | CopyOptions KeepMacros example C# | Enable macros after copying workbook Aspose.Cells | Transfer Excel VBA modules programmatically .NET
// Developer Intent: Copy VBA macros from a source workbook to a destination workbook, retain full macro functionality, and output a macro‑enabled file.
// Use Cases: Merge a template workbook containing custom VBA with a data‑driven workbook generated at runtime. | Add macros to a macro‑free workbook after exporting data so end users can run predefined scripts. | Create a consolidated report that combines multiple workbooks while preserving their original macros for further automation.
// AI Prompts: Generate C# code using Aspose.Cells to copy selected VBA modules from one workbook to another and adjust form control references. | Show how to rename or reassign form controls after copying macros with Aspose.Cells CopyOptions.KeepMacros. | Explain how to configure macro security settings when saving a workbook that contains transferred VBA macros using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopyDemo
{
    // Shows how to verify source and destination files, load them using Aspose.Cells, copy all content while preserving VBA via CopyOptions.KeepMacros, enable macros in the target workbook, and save the result as a macro‑enabled .xlsm file.
    public class MacroHelper
    {
        /// <param name="sourcePath">Path to the source macro‑enabled workbook (e.g., .xlsm).</param>
        /// <param name="destPath">Path to the destination workbook (can be macro‑free or macro‑enabled).</param>
        /// <param name="outputPath">Path where the merged workbook will be saved.</param>
        public static void CopyMacroAndUpdateControls(string sourcePath, string destPath, string outputPath)
        {
            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source file not found: {sourcePath}");

                // Verify destination file exists
                if (!File.Exists(destPath))
                    throw new FileNotFoundException($"Destination file not found: {destPath}");

                // Load workbooks
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Workbook destWorkbook = new Workbook(destPath);

                // Copy entire content, preserving macros
                CopyOptions copyOptions = new CopyOptions { KeepMacros = true };
                destWorkbook.Copy(sourceWorkbook, copyOptions);

                // Ensure macros are enabled in the destination workbook
                destWorkbook.Settings.EnableMacros = true;

                // Save as macro‑enabled workbook
                destWorkbook.Save(outputPath, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during macro copy: {ex.Message}");
                throw;
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string sourceFile = "SourceWithMacro.xlsm";
                string destinationFile = "Destination.xlsx"; // can be macro‑free
                string resultFile = "MergedWithMacro.xlsm";

                CopyMacroAndUpdateControls(sourceFile, destinationFile, resultFile);

                Console.WriteLine($"Macro copied and workbook saved to: {resultFile}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
