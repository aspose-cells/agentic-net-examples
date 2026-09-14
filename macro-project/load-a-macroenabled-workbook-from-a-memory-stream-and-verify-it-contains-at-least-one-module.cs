// Title: Load a macro‑enabled XLSM workbook from a MemoryStream and verify it contains at least one VBA module with Aspose.Cells for .NET
// AI Prompts: Create a C# method that accepts a byte array, wraps it in a MemoryStream, loads the workbook using Aspose.Cells LoadOptions.Auto, and returns true when the workbook's VbaProject includes one or more modules. | Develop a C# console application that reads an .xlsm file path, converts the file to a byte array, calls the validation method, and prints whether the workbook contains any VBA modules.
// Common Searches: aspnet load xlsm from byte array and check for VBA modules | how to programmatically verify a macro-enabled Excel file contains VBA code using Aspose.Cells | C# detect presence of VBA project in an in‑memory workbook | Aspose.Cells LoadOptions.Auto example for macro-enabled workbooks | validate that an Excel workbook loaded from MemoryStream has at least one VBA module
// Tags: load macro-enabled workbook memory stream Aspose.Cells | detect VBA modules C# Aspose.Cells | validate VBA project existence .NET | auto-detect workbook format Aspose.Cells | check XLSM for VBA code programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Demonstrates loading a macro‑enabled Excel workbook from a byte array via MemoryStream and using Aspose.Cells to confirm the workbook contains at least one VBA module.
    public class MacroWorkbookValidator
    {
        // Returns true if the workbook loaded from the given byte array contains at least one VBA module.
        public static bool HasVbaModule(byte[] workbookBytes)
        {
            if (workbookBytes == null || workbookBytes.Length == 0)
                return false;

            try
            {
                using (MemoryStream ms = new MemoryStream(workbookBytes))
                {
                    // Let Aspose.Cells auto‑detect the format (including macro‑enabled files).
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                    Workbook workbook = new Workbook(ms, loadOptions);

                    // Verify that the VBA project exists and has at least one module.
                    return workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
                return false;
            }
        }

        // Entry point for testing the validator.
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("Please provide the path to an XLSM file as an argument.");
                    return;
                }

                string filePath = args[0];

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                byte[] workbookBytes = File.ReadAllBytes(filePath);
                bool containsVba = HasVbaModule(workbookBytes);

                Console.WriteLine(containsVba
                    ? "The workbook contains at least one VBA module."
                    : "The workbook does not contain any VBA modules.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
