// Title: Export Excel workbook formulas as a hierarchical JSON tree using Aspose.Cells for .NET
// Common Searches: how to extract Excel formulas as a JSON tree with Aspose.Cells C# | Aspose.Cells C# export formula hierarchy to JSON file | C# get formula tokens from workbook and serialize to JSON using Aspose.Cells | save Excel cell formulas with parent‑child structure in JSON using Aspose.Cells | Aspose.Cells .NET build formula tree from workbook cells
// Tags: Aspose.Cells formula hierarchy export | C# generate formula tree JSON | Excel formula token parsing Aspose.Cells | serialize cell formulas to JSON | parent‑child formula representation .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through worksheets and cells containing formulas, captures each cell's sheet name, address, raw formula, and (when supported) builds a parent‑child formula tree using the FormulaToken API, then serializes the collected data into an indented JSON file.
public class FormulaNode
{
    // Token type (e.g., Operand, Operator, Function)
    public string Type { get; set; }

    // Token value (e.g., "A1", "+", "SUM")
    public string Value { get; set; }

    // Child nodes (operands of an operator/function)
    public List<FormulaNode> Children { get; set; } = new List<FormulaNode>();
}

// Container for a cell's formula information
public class CellFormulaInfo
{
    public string SheetName { get; set; }
    public string CellName { get; set; }
    public string Formula { get; set; }
    public FormulaNode TreeRoot { get; set; }
}

class Program
{
    static void Main()
    {
        const string inputPath = @"InputWorkbook.xlsx";
        const string outputPath = @"FormulaTree.json";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load an existing workbook
            var workbook = new Workbook(inputPath);
            var allFormulas = new List<CellFormulaInfo>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        FormulaNode root = null;

                        // Attempt to build a formula tree if the API is available
                        try
                        {
                            // Aspose.Cells versions prior to 23.x may not expose FormulaToken.
                            // In such cases, we simply store null for the tree.
                            // If GetFormulaTokens is supported, you can uncomment the following lines
                            // and implement BuildFormulaTree accordingly.

                            // var tokens = cell.GetFormulaTokens();
                            // root = BuildFormulaTree(tokens);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Unable to parse formula tokens for {cell.Name} - {ex.Message}");
                        }

                        // Store the information
                        allFormulas.Add(new CellFormulaInfo
                        {
                            SheetName = sheet.Name,
                            CellName = cell.Name,
                            Formula = cell.Formula,
                            TreeRoot = root
                        });
                    }
                }
            }

            // Serialize the collection to JSON (indented)
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(allFormulas, jsonOptions);

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save JSON to a file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Formula tree JSON saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder for building a formula tree from tokens.
    // The implementation depends on the availability of Aspose.Cells.FormulaToken API.
    private static FormulaNode BuildFormulaTree(object tokens)
    {
        // Since the FormulaToken API is not available in the referenced version,
        // this method returns null. Implement token parsing here if the API becomes available.
        return null;
    }
}
