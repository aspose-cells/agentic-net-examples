// Title: C# – Export each Excel worksheet to a separate TSV file using Aspose.Cells
// Description: Loads an Excel workbook with Aspose.Cells, loops through all worksheets, sets each sheet as active, configures TxtSaveOptions for SaveFormat.Tsv with ExportAllSheets disabled, and saves the active sheet to a uniquely named TSV file (including sheet index and name). The workbook is then disposed.
// Keywords: Aspose.Cells C# | export worksheet to TSV | Excel to TSV .NET | TxtSaveOptions SaveFormat.Tsv | save each sheet as TSV | multiple TSV files from workbook | Aspose.Cells SaveOptions | C# Excel conversion | TSV export Aspose | Excel multi‑sheet TSV export
// Common Searches: Aspose.Cells export each sheet to TSV | C# save Excel worksheets as separate TSV files | How to loop through worksheets and create TSV files | TxtSaveOptions Tsv example Aspose.Cells | Convert multi‑sheet Excel to multiple TSV files .NET
// Developer Intent: Generate individual TSV files for every worksheet in an Excel workbook.
// Use Cases: Create per‑sheet TSV extracts for data‑pipeline ingestion. | Produce separate TSV reports for departmental worksheets. | Automate bulk conversion of multi‑sheet Excel reports into TSV for database import.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through all worksheets in a workbook and save each one as a TSV file named with its index and sheet name. | Show how to set TxtSaveOptions for TSV format with ExportAllSheets = false to export only the active worksheet. | Explain how to modify the example to use a custom delimiter, skip hidden sheets, or write the TSV files to a specific folder.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook with Aspose.Cells, loops through all worksheets, sets each sheet as active, configures TxtSaveOptions for SaveFormat.Tsv with ExportAllSheets disabled, and saves the active sheet to a uniquely named TSV file (including sheet index and name). The workbook is then disposed.
    class ExportWorksheetsToTsv
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (lifecycle rule: Workbook(string))
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as the active sheet
                workbook.Worksheets.ActiveSheetIndex = i;

                // Create TSV save options (lifecycle rule: TxtSaveOptions(SaveFormat))
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv);

                // Export only the active sheet (rule: ExportAllSheets property)
                saveOptions.ExportAllSheets = false;

                // Build output file name using sheet index and name
                string sheetName = workbook.Worksheets[i].Name;
                string outputPath = $"Sheet{i + 1}_{sheetName}.tsv";

                // Save the active worksheet as a TSV file (lifecycle rule: Workbook.Save(string, SaveOptions))
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Worksheet \"{sheetName}\" saved to \"{outputPath}\"");
            }

            // Dispose the workbook when done
            workbook.Dispose();
        }
    }
}
