// Title: Enumerate initialized cells and flag those without a value using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enumerate cells and find initialized cells that have no value | How to detect blank but formatted cells using Cells enumerator in Aspose.Cells | Create a report of cells with IsNull type and not formulas in a .NET workbook | Flag empty cells while iterating over worksheet cells with Aspose.Cells for .NET | List cell addresses of initialized cells lacking data using Aspose.Cells C#
// Tags: Aspose.Cells iterate over initialized cells | detect cells with CellValueType.IsNull | exclude formula cells when checking for emptiness | record flagged cell addresses in report worksheet | save workbook with empty‑cell analysis Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace CellsEnumeratorFlagging
{
    // The example loads or creates an Excel workbook, uses the Cells enumerator to walk through all initialized cells in the first worksheet, flags cells whose Type is IsNull and that are not formulas, writes their addresses to a new "FlaggedCells" worksheet, and saves the workbook with the generated report.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook. Adjust as needed.
            string inputPath = "input.xlsx";

            Workbook workbook;

            // Load existing workbook or create a sample one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from '{inputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                try
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    Cells cells = ws.Cells;

                    // Cells with values
                    cells["A1"].PutValue("Header");
                    cells["B2"].PutValue(123);
                    cells["C3"].PutValue(DateTime.Now);

                    // Cells with formulas (they have a value after calculation)
                    cells["D4"].Formula = "=SUM(B2:C3)";

                    // Cells that are formatted but have no value
                    Style style = workbook.CreateStyle();
                    style.ForegroundColor = Color.Yellow;
                    style.Pattern = BackgroundType.Solid;
                    cells["E5"].SetStyle(style); // No value, just style
                    cells["F6"].SetStyle(style); // No value, just style

                    // Cells that are explicitly cleared (no style, no value)
                    cells["G7"].PutValue(string.Empty); // Clears any existing value

                    // Save the sample workbook for reference
                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook created at '{inputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            // Process the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells worksheetCells = worksheet.Cells;

            List<string> flaggedCellAddresses = new List<string>();

            // Enumerate initialized cells
            IEnumerator cellEnumerator = worksheetCells.GetEnumerator();

            try
            {
                while (cellEnumerator.MoveNext())
                {
                    Cell currentCell = cellEnumerator.Current as Cell;
                    if (currentCell == null)
                        continue;

                    // A cell without a value has Type == IsNull and is not a formula
                    bool hasNoValue = currentCell.Type == CellValueType.IsNull && !currentCell.IsFormula;

                    if (hasNoValue)
                    {
                        flaggedCellAddresses.Add(currentCell.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during cell enumeration: {ex.Message}");
            }

            // Output results to console
            Console.WriteLine("\n=== Flagged Cells (initialized but without a value) ===");
            if (flaggedCellAddresses.Count == 0)
            {
                Console.WriteLine("No such cells were found.");
            }
            else
            {
                foreach (string address in flaggedCellAddresses)
                {
                    Console.WriteLine(address);
                }
            }

            // Create or clear a report worksheet
            Worksheet reportSheet = workbook.Worksheets["FlaggedCells"];
            if (reportSheet == null)
            {
                reportSheet = workbook.Worksheets.Add("FlaggedCells");
            }
            else
            {
                reportSheet.Cells.Clear();
            }

            // Write report
            reportSheet.Cells[0, 0].PutValue("Flagged Cells (Initialized but No Value)");
            for (int i = 0; i < flaggedCellAddresses.Count; i++)
            {
                reportSheet.Cells[i + 1, 0].PutValue(flaggedCellAddresses[i]);
            }

            // Save the workbook with the report
            string outputPath = "output_flagged.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"\nReport saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
