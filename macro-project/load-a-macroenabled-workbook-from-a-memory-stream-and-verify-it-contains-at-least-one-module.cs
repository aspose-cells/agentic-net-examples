// Title: Load XLSM from MemoryStream and verify VBA modules using Aspose.Cells (C#)
// Description: Demonstrates how to read a macro‑enabled .xlsm file into a byte array, create a MemoryStream, load it with Aspose.Cells, check the Workbook.HasMacro flag, and confirm that the VbaProject contains at least one VBA module.
// Keywords: Aspose.Cells C# load xlsm | memory stream workbook Aspose | Workbook.HasMacro check | VbaProject modules count | macro‑enabled workbook verification | read xlsm from byte array | detect VBA modules Aspose.Cells | security scan macro files .NET | Aspose.Cells VBA project inspection | load macro workbook from stream
// Common Searches: How to open an xlsm file from a MemoryStream with Aspose.Cells | Check if a workbook has macros using Aspose.Cells C# | Get VBA module count from a macro‑enabled workbook in .NET | Aspose.Cells load macro workbook from byte array | Verify presence of VBA modules in an XLSM file
// Developer Intent: Load a macro‑enabled workbook from a MemoryStream and ensure it contains at least one VBA module.
// Use Cases: Validate uploaded XLSM files on a web server before processing them. | Confirm that template workbooks include required VBA modules for automated reporting. | Perform a security audit by detecting macros and VBA modules in user‑submitted spreadsheets.
// AI Prompts: Generate C# code that opens an .xlsm file from a byte array, checks Workbook.HasMacro, and lists all VBA module names with Aspose.Cells. | Explain step‑by‑step how to load a macro‑enabled workbook from a MemoryStream and verify the VbaProject contains at least one module using Aspose.Cells for .NET. | Provide best‑practice error handling for reading macro‑enabled workbooks from streams and detecting missing VBA modules with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCheck
{
    // Demonstrates how to read a macro‑enabled .xlsm file into a byte array, create a MemoryStream, load it with Aspose.Cells, check the Workbook.HasMacro flag, and confirm that the VbaProject contains at least one VBA module.
    public class MacroVerification
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to a macro‑enabled workbook (xlsm). Adjust as needed.
            string macroWorkbookPath = "sample_with_macro.xlsm";

            // Ensure the file exists to avoid FileNotFoundException.
            if (!File.Exists(macroWorkbookPath))
            {
                Console.WriteLine($"File not found: {macroWorkbookPath}");
                return;
            }

            try
            {
                // Load the file into a memory stream.
                byte[] fileBytes = File.ReadAllBytes(macroWorkbookPath);
                using (MemoryStream memoryStream = new MemoryStream(fileBytes))
                {
                    // Reset position to the beginning before loading.
                    memoryStream.Position = 0;

                    // Load the workbook from the memory stream.
                    Workbook workbook = new Workbook(memoryStream);

                    // Verify that the workbook reports having macros.
                    bool hasMacro = workbook.HasMacro;
                    Console.WriteLine($"Workbook.HasMacro: {hasMacro}");

                    // Verify that the VBA project contains at least one module.
                    bool hasModules = false;
                    if (workbook.VbaProject != null && workbook.VbaProject.Modules != null)
                    {
                        hasModules = workbook.VbaProject.Modules.Count > 0;
                    }
                    Console.WriteLine($"Workbook contains at least one VBA module: {hasModules}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
