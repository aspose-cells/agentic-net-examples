// Title: Copy VBA macros between .xlsm workbooks with Aspose.Cells for .NET
// Description: Demonstrates how to load a macro‑enabled workbook, create a new workbook, transfer the EnableMacros security flag, and copy all VBA macros using CopyOptions.KeepMacros. The example also shows how to generate a minimal .xlsm file when the source is missing and saves the result as a macro‑enabled workbook.
// Keywords: Aspose.Cells copy macros | C# copy VBA macros | KeepMacros option | EnableMacros setting | macro‑enabled workbook .NET | copy workbook with macros | Aspose.Cells Xlsm
// Common Searches: how to copy macros with Aspose.Cells | preserve macro security when copying .xlsm files | Aspose.Cells CopyOptions KeepMacros example | C# duplicate macro‑enabled workbook | create minimal .xlsm workbook programmatically
// Developer Intent: Transfer all VBA macros from an existing .xlsm file to a new workbook while keeping the original macro security configuration.
// Use Cases: Generate report workbooks that inherit macros from a template without losing VBA code. | Automate bulk duplication of macro‑enabled files across folders while maintaining security settings. | Create a fresh macro‑enabled workbook on‑the‑fly when the source template is unavailable.
// AI Prompts: Write C# code using Aspose.Cells to copy all macros from source.xlsm to destination.xlsm and preserve the EnableMacros flag. | Show how to use CopyOptions.KeepMacros to duplicate a macro‑enabled workbook in .NET. | Explain how to handle a missing source .xlsm by programmatically creating a minimal macro‑enabled workbook before copying macros.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to load a macro‑enabled workbook, create a new workbook, transfer the EnableMacros security flag, and copy all VBA macros using CopyOptions.KeepMacros. The example also shows how to generate a minimal .xlsm file when the source is missing and saves the result as a macro‑enabled workbook.
    public class CopyMacrosDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Macro copy operation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string sourcePath = "source_with_macros.xlsm";
            const string destPath = "destination_with_copied_macros.xlsm";

            // Ensure the source workbook exists; create a minimal macro‑enabled workbook if missing
            if (!File.Exists(sourcePath))
            {
                Workbook temp = new Workbook();
                // Enable macros flag for the new workbook
                temp.Settings.EnableMacros = true;
                temp.Save(sourcePath, SaveFormat.Xlsm);
            }

            // Load the source workbook that contains macros
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // Preserve the macro security setting from the source workbook
            destWorkbook.Settings.EnableMacros = sourceWorkbook.Settings.EnableMacros;

            // Set copy options to keep macros during the copy operation
            CopyOptions options = new CopyOptions
            {
                KeepMacros = true
            };

            // Copy the entire source workbook (including macros) into the destination workbook
            sourceWorkbook.Copy(destWorkbook, options);

            // Save the destination workbook as a macro‑enabled file
            destWorkbook.Save(destPath, SaveFormat.Xlsm);
        }
    }
}
