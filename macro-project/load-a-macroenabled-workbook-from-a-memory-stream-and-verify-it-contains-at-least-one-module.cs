// Title: Load a macro‑enabled .xlsm workbook from a MemoryStream and verify VBA modules using Aspose.Cells for .NET
// Common Searches: C# load .xlsm file from memory stream using Aspose.Cells | How to determine if a macro‑enabled workbook contains VBA modules with Aspose.Cells .NET | Using LoadOptions Xlsx to open a macro workbook from a byte array in C# | Check for VBA modules in an Aspose.Cells workbook loaded from a stream
// Tags: load xlsm workbook from memorystream c# | verify vba modules aspocells | loadoptions xlsx macro workbook | vba project modules count aspocells | read macro-enabled excel bytes c#

using System;
using System.IO;
using Aspose.Cells;

// The example reads a macro‑enabled Excel file (.xlsm) into a byte array, wraps it in a MemoryStream, and loads it with Aspose.Cells using LoadOptions set to Xlsx. It then inspects the Workbook.VbaProject to confirm that at least one VBA module exists, outputting the result while handling missing files and other exceptions.
class MacroWorkbookVerifier
{
    static void Main()
    {
        try
        {
            const string filePath = "SampleMacroWorkbook.xlsm";

            // Ensure the source file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            // Read the workbook bytes.
            byte[] workbookBytes = File.ReadAllBytes(filePath);

            // Load the workbook from a memory stream. Use Xlsx format for compatibility with older Aspose.Cells versions.
            using (MemoryStream ms = new MemoryStream(workbookBytes))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(ms, loadOptions);

                // Verify that the workbook contains at least one VBA module.
                bool hasModules = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;

                Console.WriteLine(hasModules
                    ? "The workbook contains at least one VBA module."
                    : "The workbook does NOT contain any VBA modules.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
