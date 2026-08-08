// Title: C# – Remove Macros from XLSM and Save as Strict Open XML XLSX with Aspose.Cells
// Description: Loads a macro‑enabled XLSM workbook using Aspose.Cells, calls Workbook.RemoveMacro() to strip all VBA code, and saves the clean file as a strict Open XML XLSX. Includes file‑existence checks and exception handling for robust .NET automation.
// Keywords: Aspose.Cells | C# macro removal | Workbook.RemoveMacro | XLSM to XLSX conversion | strict Open XML | remove VBA | save as Xlsx | .NET Excel automation
// Common Searches: How to delete VBA macros from XLSM using Aspose.Cells C# | Convert macro‑enabled Excel file to macro‑free XLSX .NET | Aspose.Cells remove macro and save strict Open XML | C# code to strip macros from workbook | Batch remove macros from XLSM files Aspose.Cells
// Developer Intent: Strip all VBA macros from a macro‑enabled Excel workbook and output a strict Open XML XLSX file with Aspose.Cells for .NET.
// Use Cases: Sanitize user‑uploaded spreadsheets by removing macros before processing or storage. | Batch‑convert legacy macro‑enabled reports to macro‑free XLSX for archiving or compliance. | Prepare workbooks for environments that only accept strict Open XML formats, ensuring no VBA code remains.
// AI Prompts: Generate C# code that opens an XLSM file with Aspose.Cells, removes all macros, and saves it as a strict Open XML XLSX, including error handling for missing files. | Explain the behavior of Workbook.RemoveMacro in Aspose.Cells and list the file formats that can be saved after macro removal. | Provide a step‑by‑step guide to iterate over a folder of XLSM files, remove macros from each, and save the results as XLSX using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroRemoval
{
    // Loads a macro‑enabled XLSM workbook using Aspose.Cells, calls Workbook.RemoveMacro() to strip all VBA code, and saves the clean file as a strict Open XML XLSX. Includes file‑existence checks and exception handling for robust .NET automation.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source macro‑enabled workbook (XLSM)
            string sourcePath = "source_with_macros.xlsm";

            // Path for the macro‑free workbook (XLSX)
            string destinationPath = "macro_free.xlsx";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook from the XLSM file
                Workbook workbook = new Workbook(sourcePath);

                // Remove all VBA/macros from the workbook
                workbook.RemoveMacro();

                // Save the cleaned workbook as a strict Open XML XLSX file
                workbook.Save(destinationPath, SaveFormat.Xlsx);

                Console.WriteLine($"Macros removed and workbook saved to: {destinationPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
