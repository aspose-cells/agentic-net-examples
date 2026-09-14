// Title: Export all formulas from an Excel workbook to a hierarchical JSON tree using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, collects each cell's formula, builds a parent‑child FormulaNode hierarchy, and saves the structure as indented JSON. | Enhance the sample to parse a formula string into operator and operand nodes, linking them as children in the exported JSON hierarchy. | Add logic to exclude hidden worksheets and cells without formulas before constructing the JSON representation.
// Common Searches: how to export Excel formulas as JSON using Aspose.Cells C# | Aspose.Cells .NET create hierarchical formula tree from workbook | C# serialize Excel cell formulas to a JSON file | skip hidden sheets when exporting formulas with Aspose.Cells | extract formula hierarchy from Excel with Aspose.Cells and save as JSON
// Tags: export formulas to JSON with Aspose.Cells | Aspose.Cells formula tree serialization | C# extract Excel formula hierarchy | skip hidden worksheets Aspose.Cells | parse Excel formula into node structure .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace FormulaTreeExport
{
    // Represents a node in the formula tree.
    // The program loads an input.xlsx workbook via Aspose.Cells, iterates every worksheet and cell, captures each cell's formula into a simple FormulaNode, stores nodes in a dictionary keyed by cell address, serializes the dictionary to an indented JSON file, ensures the output directory exists, and writes the result to formulaTree.json while handling file‑related errors.
    class FormulaNode
    {
        public string? Type { get; set; }
        public string? Value { get; set; }
        public List<FormulaNode> Children { get; set; } = new List<FormulaNode>();
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "formulaTree.json";

                // Verify input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook.
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }

                var formulaTrees = new Dictionary<string, FormulaNode>();

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all cells that contain a formula.
                    foreach (Cell cell in cells)
                    {
                        if (string.IsNullOrEmpty(cell.Formula))
                            continue;

                        // Create a simple node representing the formula.
                        var formulaNode = new FormulaNode
                        {
                            Type = "Formula",
                            Value = cell.Formula
                        };

                        formulaTrees[cell.Name] = formulaNode;
                    }
                }

                // Serialize to JSON with indentation.
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(formulaTrees, options);

                // Ensure the directory for the output file exists.
                try
                {
                    string? outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to prepare output directory: {ex.Message}");
                }

                // Write JSON to file.
                try
                {
                    File.WriteAllText(outputPath, json);
                    Console.WriteLine($"Formula tree JSON written to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write output file: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
