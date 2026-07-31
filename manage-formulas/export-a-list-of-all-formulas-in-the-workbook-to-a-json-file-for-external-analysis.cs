// Title: Export All Excel Formulas to JSON with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, parses every worksheet, gathers non‑empty formulas, writes them to a single column, creates a range and uses Aspose.Cells JsonSaveOptions with JsonUtility.ExportRangeToJson to produce a JSON file for external analysis.
// Keywords: Aspose.Cells | C# | Export formulas to JSON | Excel formula extraction | JsonUtility | JsonSaveOptions | Workbook parsing | Range export | Spreadsheet auditing | Data migration
// Common Searches: Aspose.Cells export formulas as JSON | C# extract all Excel formulas | JsonUtility ExportRangeToJson example | How to save Excel formulas to a JSON file | Parse formulas with Aspose.Cells .NET
// Developer Intent: Generate a JSON file containing every formula from an Excel workbook for downstream processing.
// Use Cases: Create an inventory of workbook calculations for audit trails. | Feed extracted formulas into a custom validation engine outside Excel. | Produce documentation of spreadsheet logic in a machine‑readable format.
// AI Prompts: Write a C# method that returns a JSON string of all formulas in a workbook using Aspose.Cells. | Show how to include the worksheet name alongside each formula in the exported JSON. | Suggest performance optimizations for exporting formulas from very large workbooks to JSON.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsFormulaExport
{
    // Loads an Excel workbook, parses every worksheet, gathers non‑empty formulas, writes them to a single column, creates a range and uses Aspose.Cells JsonSaveOptions with JsonUtility.ExportRangeToJson to produce a JSON file for external analysis.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "input.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook (create/load lifecycle rule)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Ensure all formulas are parsed
                sourceWorkbook.ParseFormulas(false);

                // Collect all formulas from every worksheet
                List<string> formulaList = new List<string>();
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // Iterate through used cells only for efficiency
                    foreach (Cell cell in cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            formulaList.Add(cell.Formula);
                        }
                    }
                }

                // Create a new workbook to hold the list of formulas (create lifecycle rule)
                Workbook exportWorkbook = new Workbook();
                Worksheet exportSheet = exportWorkbook.Worksheets[0];

                // Write each formula into column A of the export sheet
                for (int i = 0; i < formulaList.Count; i++)
                {
                    exportSheet.Cells[i, 0].PutValue(formulaList[i]);
                }

                // Define the range that contains the formulas (use fully qualified Aspose.Cells.Range)
                Aspose.Cells.Range exportRange = exportSheet.Cells.CreateRange(0, 0, formulaList.Count, 1);

                // Configure JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    ExportAsString = true,   // export cell values as strings
                    HasHeaderRow = false,    // no header row needed
                    ExportEmptyCells = false // skip empty cells
                };

                // Export the range to a JSON string using the provided rule
                string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Save the JSON string to a file (save lifecycle rule)
                string outputPath = "formulas.json";
                File.WriteAllText(outputPath, jsonOutput);

                Console.WriteLine($"Exported {formulaList.Count} formulas to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
