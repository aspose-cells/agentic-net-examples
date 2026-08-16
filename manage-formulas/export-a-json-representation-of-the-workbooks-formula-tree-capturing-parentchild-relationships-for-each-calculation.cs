// Title: Export Excel Formula Dependency Tree as Nested JSON with Aspose.Cells for .NET
// Description: Creates a workbook, adds numeric values and inter‑dependent formulas, calculates them, and saves the file as nested JSON using Aspose.Cells JsonSaveOptions (ExportNestedStructure). The output captures parent‑child relationships of each formula.
// Keywords: Aspose.Cells | C# | .NET | Export formula tree JSON | nested JSON formula hierarchy | JsonSaveOptions ExportNestedStructure | Excel formula dependency | parent child formula export
// Common Searches: Aspose.Cells export formula tree to JSON | how to get Excel formula dependencies as JSON in C# | JsonSaveOptions nested structure example | export Excel formulas with parent child hierarchy | C# Aspose.Cells formula dependency export
// Developer Intent: Generate a JSON file that represents the workbook’s formula dependency tree with explicit parent‑child links.
// Use Cases: Document and audit formula relationships for compliance or review. | Feed a custom calculation engine that requires explicit cell dependency data. | Render an interactive formula hierarchy in a web UI using the exported JSON.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook’s formula tree as nested JSON, preserving parent‑child links. | Explain how ExportNestedStructure and AlwaysExportAsJsonObject affect the JSON output of formulas. | Show how to parse the generated FormulaTree.json in JavaScript to build a visual dependency graph.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTreeExport
{
    // Creates a workbook, adds numeric values and inter‑dependent formulas, calculates them, and saves the file as nested JSON using Aspose.Cells JsonSaveOptions (ExportNestedStructure). The output captures parent‑child relationships of each formula.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add formulas that depend on each other to form a calculation tree
            // B1 depends on A1 and A2
            cells["B1"].Formula = "=A1+A2";
            // B2 depends on A2 and A3
            cells["B2"].Formula = "=A2+A3";
            // C1 depends on B1 and B2 (parent node)
            cells["C1"].Formula = "=B1+B2";

            // Ensure formulas are calculated (optional, but useful for verification)
            workbook.CalculateFormula();

            // Configure JSON save options to export as a nested (parent‑child) structure
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportNestedStructure = true,          // Enable parent‑child hierarchy
                AlwaysExportAsJsonObject = true,      // Export workbook as a JSON object even if single sheet
                ExportEmptyCells = false,              // Skip empty cells for cleaner output
                HasHeaderRow = false,                  // No header row needed for formula tree
                ExportAsString = true                  // Export cell values as strings
            };

            // Save the workbook as JSON; the resulting file contains the formula tree
            string outputPath = "FormulaTree.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Formula tree exported to: {outputPath}");
        }
    }
}
