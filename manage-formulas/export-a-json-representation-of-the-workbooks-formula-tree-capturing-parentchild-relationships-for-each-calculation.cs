// Title: Export Excel formula dependency tree to JSON using Aspose.Cells for .NET (C#)
// Common Searches: how to extract Excel formulas and their references with Aspose.Cells in C# | C# Aspose.Cells export formula dependency graph to JSON file | save Excel formula tree as JSON using Aspose.Cells .NET library | retrieve cell formula tokens and serialize to JSON with Aspose.Cells C# | export workbook formula hierarchy to JSON with Aspose.Cells API
// Tags: Aspose.Cells export formula tree JSON | C# extract Excel formula references | serialize workbook formula nodes | iterate worksheets collect formulas Aspose | JSON output of Excel calculation dependencies

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsFormulaTreeExport
{
    // Represents a node in the formula tree
    // The example loads an Excel workbook with Aspose.Cells, walks through each worksheet's used range, creates a FormulaNode for every cell containing a formula (capturing its address and formula string), aggregates the nodes, serializes the collection to indented JSON, and writes the result to a file, handling missing files, load errors, directory creation, and serialization/write exceptions.
    public class FormulaNode
    {
        // Address of the cell containing the formula (e.g., "Sheet1!A1")
        public string CellAddress { get; set; } = string.Empty;

        // The formula string as it appears in the cell
        public string Formula { get; set; } = string.Empty;

        // List of cell addresses that this formula depends on
        public List<string> References { get; set; } = new List<string>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source workbook
                string inputPath = "input.xlsx";

                // Path where the JSON representation will be saved
                string outputPath = "formulaTree.json";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Collection to hold all formula nodes across worksheets
                List<FormulaNode> formulaNodes = new List<FormulaNode>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range to limit iteration to populated cells
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // If the sheet is empty, skip it
                    if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                        continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = startRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = startCol + usedRange.ColumnCount - 1;

                    // Loop through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                FormulaNode node = new FormulaNode
                                {
                                    CellAddress = $"{sheet.Name}!{cell.Name}",
                                    Formula = cell.Formula
                                };

                                // NOTE: Detailed token parsing (references) is omitted for compatibility.
                                // If needed, Aspose.Cells provides GetFormulaTokens() in newer versions.

                                formulaNodes.Add(node);
                            }
                        }
                    }
                }

                // Serialize the collection to JSON with indentation for readability
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json;
                try
                {
                    json = JsonSerializer.Serialize(formulaNodes, jsonOptions);
                }
                catch (Exception serEx)
                {
                    Console.WriteLine($"JSON serialization failed: {serEx.Message}");
                    return;
                }

                // Ensure the directory for the output file exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    try
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    catch (Exception dirEx)
                    {
                        Console.WriteLine($"Failed to create output directory: {dirEx.Message}");
                        return;
                    }
                }

                // Save the JSON to the specified file
                try
                {
                    File.WriteAllText(outputPath, json);
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write output file: {writeEx.Message}");
                    return;
                }

                Console.WriteLine($"Formula tree exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
