// Title: Remove VBA macros from an XLSM workbook and save it as a strict OpenXML XLSX using Aspose.Cells for .NET
// AI Prompts: Load a macro‑enabled XLSM file with Aspose.Cells, call Workbook.RemoveMacro to strip all VBA code, and then save the workbook as a strict OpenXML XLSX. | Write C# that verifies the source XLSM exists, removes its macros, and uses SaveFormat.Xlsx to export a macro‑free file.
// Common Searches: Aspose.Cells C# remove VBA macros from XLSM and save as strict XLSX | How to convert a macro‑enabled Excel file to a macro‑free strict OpenXML format using Aspose.Cells | C# code to load an XLSM, delete macros and export to XLSX with SaveFormat.Xlsx | Remove macros from an Excel workbook programmatically with Aspose.Cells .NET | Save workbook as strict OpenXML XLSX after stripping VBA in ASP.NET
// Tags: remove VBA macros Aspose.Cells | convert XLSM to strict OpenXML XLSX | save workbook as strict XLSX C# | load macro‑enabled workbook Aspose.Cells | Workbook.RemoveMacro usage

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example demonstrates loading a macro‑enabled XLSM workbook, checking the source file, removing all VBA macros via Workbook.RemoveMacro, and saving the result as a strict OpenXML XLSX file using SaveFormat.Xlsx.
    public class RemoveMacroAndSaveStrictXlsx
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the macro-enabled source file
            string sourcePath = "input.xlsm";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook (macro-enabled)
            Workbook workbook = new Workbook(sourcePath);

            // Remove any VBA macros from the workbook
            workbook.RemoveMacro();

            // Save the workbook as a macro‑free strict Open XML XLSX file
            string destPath = "output.xlsx";

            // Save with Xlsx format (strict Open XML)
            workbook.Save(destPath, SaveFormat.Xlsx);

            Console.WriteLine($"Macros removed and saved to {destPath}");
        }
    }
}
