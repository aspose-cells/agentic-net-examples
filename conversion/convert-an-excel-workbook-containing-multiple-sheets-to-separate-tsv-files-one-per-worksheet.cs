// Title: Convert each worksheet in an Excel workbook to separate TSV files using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, loops through all worksheets, and saves each one as an individual .tsv file using TxtSaveOptions. | Show how to configure TxtSaveOptions with SaveFormat.Tsv to export only the active worksheet inside a worksheet‑iteration loop.
// Common Searches: Aspose.Cells C# export each sheet of an Excel file to a separate TSV file | How to save individual worksheets as tab‑separated values with Aspose.Cells .NET | Programmatically split a multi‑sheet Excel workbook into multiple .tsv files using C# | Loop through worksheets in Aspose.Cells and generate TSV files per sheet | Export active worksheet to TSV using TxtSaveOptions in Aspose.Cells
// Tags: Aspose.Cells worksheet to TSV conversion | TxtSaveOptions SaveFormat.Tsv usage | export active sheet as tab‑separated file | split Excel workbook into multiple TSV files C# | iterate workbook worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, iterates over its worksheets, sets each sheet as the active one, and saves it as a separate TSV file using TxtSaveOptions with ExportAllSheets disabled. Output files are named with the original workbook name and the worksheet name.
    public class WorkbookToSeparateTsv
    {
        public static void Run()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook from file
                using (Workbook workbook = new Workbook(sourcePath))
                {
                    // Iterate through each worksheet in the workbook
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        // Set the current worksheet as the active sheet
                        workbook.Worksheets.ActiveSheetIndex = i;

                        // Prepare TSV save options: export only the active sheet
                        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
                        {
                            ExportAllSheets = false // ensure only the active sheet is saved
                        };

                        // Build output file name: original name + sheet name + .tsv
                        string sheetName = workbook.Worksheets[i].Name;
                        string outputFile = Path.Combine(
                            Path.GetDirectoryName(sourcePath) ?? string.Empty,
                            $"{Path.GetFileNameWithoutExtension(sourcePath)}_{sheetName}.tsv");

                        // Save the active sheet as a TSV file
                        workbook.Save(outputFile, saveOptions);

                        Console.WriteLine($"Saved sheet '{sheetName}' to '{outputFile}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToSeparateTsv.Run();
        }
    }
}
