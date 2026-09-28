// Title: Create and save a macro‑enabled .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Instantiate a Workbook object and persist it as a macro‑enabled .xlsm file with Aspose.Cells. | Use the SaveFormat.Xlsm enum to generate a macro‑enabled Excel workbook in C#. | Wrap the workbook creation and save operations in a try‑catch block to handle Aspose.Cells errors.
// Common Searches: asp.net create macro enabled xlsm workbook with Aspose.Cells | save workbook as xlsm using Aspose.Cells C# example | how to generate a macro enabled Excel file programmatically with Aspose.Cells | Aspose.Cells SaveFormat.Xlsm usage in .NET
// Tags: create workbook Aspose.Cells | save workbook as xlsm Aspose.Cells | macro enabled Excel file Aspose.Cells | Aspose.Cells SaveFormat.Xlsm | exception handling Aspose.Cells workbook

using System;
using Aspose.Cells;

namespace AsposeCellsMacroExample
{
    // The sample creates a new Workbook, specifies an output path "MacroEnabled.xlsm", saves the workbook as a macro‑enabled .xlsm file using SaveFormat.Xlsm, and reports success or any exception to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default format is .xlsx)
                Workbook workbook = new Workbook();

                // Define output path
                string outputPath = "MacroEnabled.xlsm";

                // Save the workbook as a macro‑enabled file
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
