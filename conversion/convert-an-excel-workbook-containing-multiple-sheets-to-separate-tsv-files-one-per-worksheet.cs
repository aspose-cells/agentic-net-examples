// Title: C# – Export Excel worksheets to individual TSV files with Aspose.Cells
// Description: Loads an Excel workbook, creates an output folder, iterates through every worksheet, sets it as active, and saves each one as a separate Tab‑Separated Values (TSV) file using TxtSaveOptions (ExportAllSheets = false).
// Keywords: Aspose.Cells C# export TSV | save Excel sheet as tab delimited | convert multi‑sheet workbook to TSV | TxtSaveOptions SaveFormat.Tsv example | C# Excel to TSV batch conversion
// Common Searches: Aspose.Cells export each worksheet to separate TSV | C# code to save Excel sheets as .tsv files | how to use TxtSaveOptions ExportAllSheets false | convert Excel workbook to multiple TSV files .NET | Aspose.Cells TSV output per sheet
// Developer Intent: Generate one TSV file per worksheet from an Excel workbook.
// Use Cases: Produce department‑specific TSV reports from a master workbook. | Prepare separate tab‑delimited files for bulk database imports. | Automate pipeline steps that require individual sheet exports for downstream analytics.
// AI Prompts: Write C# code using Aspose.Cells that converts every worksheet in an Excel file into separate TSV files, ensuring only the active sheet is saved each time. | Explain the role of TxtSaveOptions, SaveFormat.Tsv, and ExportAllSheets = false when exporting worksheets individually. | Suggest a robust method to sanitize worksheet names for valid file‑system names during TSV file creation.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, creates an output folder, iterates through every worksheet, sets it as active, and saves each one as a separate Tab‑Separated Values (TSV) file using TxtSaveOptions (ExportAllSheets = false).
    public class WorkbookToSeparateTsv
    {
        public static void Run()
        {
            try
            {
                // Path to the source Excel workbook
                string sourcePath = "input.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Directory where TSV files will be saved
                string outputDir = "tsv_output";
                Directory.CreateDirectory(outputDir);

                // Load the workbook (lifecycle: create/load)
                Workbook workbook = new Workbook(sourcePath);

                // Iterate through each worksheet
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // Set the current worksheet as the active sheet
                    workbook.Worksheets.ActiveSheetIndex = i;

                    // Prepare the output file name (one TSV per sheet)
                    string sheetName = workbook.Worksheets[i].Name;
                    string tsvPath = Path.Combine(outputDir, $"{sheetName}.tsv");

                    // Configure save options for TSV (ExportAllSheets = false by default)
                    TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
                    {
                        ExportAllSheets = false // ensure only the active sheet is exported
                    };

                    // Save the active sheet as a TSV file (lifecycle: save)
                    workbook.Save(tsvPath, saveOptions);
                }

                Console.WriteLine("All worksheets have been exported to separate TSV files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToSeparateTsv.Run();
        }
    }
}
