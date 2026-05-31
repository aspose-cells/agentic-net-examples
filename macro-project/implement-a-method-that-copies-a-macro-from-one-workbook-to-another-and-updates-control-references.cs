using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopyDemo
{
    public class MacroCopier
    {
        /// <summary>
        /// Copies macros (VBA project) from a source workbook to a destination workbook,
        /// copies VBA references, enables macros in the destination and saves the result.
        /// </summary>
        /// <param name="sourcePath">Path to the macro‑enabled source workbook (e.g., .xlsm).</param>
        /// <param name="destPath">Path to the destination workbook (can be a regular .xlsx file).</param>
        /// <param name="outputPath">Path where the resulting workbook will be saved (should be .xlsm to retain macros).</param>
        public static void CopyMacroAndUpdateControls(string sourcePath, string destPath, string outputPath)
        {
            try
            {
                // Verify source file exists.
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException("Source workbook not found.", sourcePath);

                // Load the source workbook that contains macros.
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Load or create the destination workbook.
                Workbook destWorkbook;
                if (File.Exists(destPath))
                {
                    destWorkbook = new Workbook(destPath);
                }
                else
                {
                    destWorkbook = new Workbook(); // creates a new workbook with a default sheet
                }

                // -----------------------------------------------------------------
                // 1. Copy the VBA project (modules, code, etc.) from source to destination.
                // -----------------------------------------------------------------
                VbaProject sourceVba = sourceWorkbook.VbaProject;
                VbaProject destVba = destWorkbook.VbaProject;

                // Copy the entire VBA project.
                destVba.Copy(sourceVba);

                // -----------------------------------------------------------------
                // 2. Copy VBA references (external libraries) if any.
                // -----------------------------------------------------------------
                VbaProjectReferenceCollection sourceRefs = sourceVba.References;
                VbaProjectReferenceCollection destRefs = destVba.References;

                destRefs.Clear();
                destRefs.Copy(sourceRefs);

                // -----------------------------------------------------------------
                // 3. Enable macros in the destination workbook settings.
                // -----------------------------------------------------------------
                destWorkbook.Settings.EnableMacros = true;

                // -----------------------------------------------------------------
                // 4. Copy worksheets while preserving macros.
                //    KeepMacros ensures any sheet‑level macro code is retained.
                // -----------------------------------------------------------------
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                foreach (Worksheet srcSheet in sourceWorkbook.Worksheets)
                {
                    // Ensure a unique sheet name in the destination workbook.
                    string newName = srcSheet.Name;
                    int suffix = 1;
                    while (destWorkbook.Worksheets.Any(ws => ws.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
                    {
                        newName = $"{srcSheet.Name}_{suffix}";
                        suffix++;
                    }

                    // Add a new sheet with the unique name and copy contents.
                    Worksheet destSheet = destWorkbook.Worksheets.Add(newName);
                    destSheet.Copy(srcSheet, copyOptions);
                }

                // -----------------------------------------------------------------
                // 5. Save the resulting workbook as a macro‑enabled file.
                // -----------------------------------------------------------------
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                destWorkbook.Save(outputPath, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during macro copy: {ex.Message}");
                // Optionally rethrow or handle further.
                throw;
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string sourceFile = "SourceWithMacro.xlsm";
                string destinationFile = "EmptyWorkbook.xlsx"; // can be an existing file or a new one
                string resultFile = "ResultWithMacro.xlsm";

                CopyMacroAndUpdateControls(sourceFile, destinationFile, resultFile);

                Console.WriteLine($"Macro copied and saved to '{resultFile}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}