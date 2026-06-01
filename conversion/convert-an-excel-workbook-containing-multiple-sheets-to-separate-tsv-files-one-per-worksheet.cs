using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class WorkbookToSeparateTsv
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
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook
                workbook = new Workbook(sourcePath);

                // Iterate through each worksheet in the workbook
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // Set the current worksheet as the active sheet
                    workbook.Worksheets.ActiveSheetIndex = i;

                    // Configure TSV save options to export only the active sheet
                    TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
                    {
                        ExportAllSheets = false // export only active sheet
                    };

                    // Build output file name (e.g., Sheet1_SheetName.tsv)
                    string sheetName = workbook.Worksheets[i].Name;
                    string outputPath = $"Sheet{i + 1}_{sheetName}.tsv";

                    try
                    {
                        // Save the active worksheet as a TSV file
                        workbook.Save(outputPath, saveOptions);
                        Console.WriteLine($"Saved: {outputPath}");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"Failed to save sheet '{sheetName}': {saveEx.Message}");
                    }
                }
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
            }
            finally
            {
                // Clean up resources
                workbook?.Dispose();
            }
        }
    }
}