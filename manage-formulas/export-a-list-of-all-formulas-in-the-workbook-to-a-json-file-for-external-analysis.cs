using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsFormulaExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourcePath = "input.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the workbook (load rule)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new workbook to hold the formula list (create rule)
                Workbook formulaWorkbook = new Workbook();
                Worksheet sheet = formulaWorkbook.Worksheets[0];

                // Write header row
                sheet.Cells["A1"].PutValue("Sheet");
                sheet.Cells["B1"].PutValue("Cell");
                sheet.Cells["C1"].PutValue("Formula");

                int currentRow = 1; // zero‑based index, row 1 is the second row (after header)

                // Iterate through all worksheets and cells to collect formulas
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    Cells cells = ws.Cells;
                    // Scan the used range of the worksheet
                    for (int row = cells.MinDataRow; row <= cells.MaxDataRow; row++)
                    {
                        for (int col = cells.MinDataColumn; col <= cells.MaxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Populate the formula list workbook
                                sheet.Cells[currentRow, 0].PutValue(ws.Name);      // Sheet name
                                sheet.Cells[currentRow, 1].PutValue(cell.Name);   // Cell address (e.g., A1)
                                sheet.Cells[currentRow, 2].PutValue(cell.Formula);// Formula text
                                currentRow++;
                            }
                        }
                    }
                }

                // Define the range that contains the collected data
                int totalRows = currentRow; // includes header row
                Aspose.Cells.Range exportRange = sheet.Cells.CreateRange(0, 0, totalRows, 3);

                // Configure JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // First row is header
                    ExportAsString = true,        // Export cell values as strings
                    ExportEmptyCells = false,     // Skip empty cells
                    Indent = "  "                 // Pretty‑print JSON
                };

                // Export the range to a JSON string (export rule)
                string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Save the JSON string to a file (save rule)
                string jsonPath = "formulas.json";
                File.WriteAllText(jsonPath, jsonOutput);

                Console.WriteLine($"Exported {currentRow - 1} formulas to '{jsonPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}