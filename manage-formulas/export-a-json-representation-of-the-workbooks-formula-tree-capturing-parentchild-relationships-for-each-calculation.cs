// Title: Export Excel Formula Tree to Hierarchical JSON with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, configures JsonSaveOptions (ExportNestedStructure, AlwaysExportAsJsonObject, ExportAsString, ExportEmptyCells, HasHeaderRow, Indent) and saves the file as JSON. The output contains each cell’s address, value, and formula, forming a parent‑child representation of the workbook’s calculation graph.
// Keywords: Aspose.Cells | C# | .NET | export formula tree | nested JSON | ExportNestedStructure | Excel to JSON | formula dependency hierarchy | workbook JsonSaveOptions | Excel model auditing
// Common Searches: Aspose.Cells export formula hierarchy JSON C# | ExportNestedStructure option example | How to get Excel formula dependency tree with Aspose.Cells | Save workbook as hierarchical JSON .NET | C# code to export Excel formulas to JSON
// Developer Intent: Generate a JSON file that captures the workbook’s formula dependencies as a parent‑child hierarchy.
// Use Cases: Create a portable snapshot of calculation logic for analytics pipelines. | Feed a visualization service that displays formula flow charts. | Archive Excel models for version control, auditing, or documentation.
// AI Prompts: Write C# code using Aspose.Cells to export an Excel workbook’s formula tree to a nested JSON file, preserving cell addresses and formulas. | Explain how the ExportNestedStructure flag builds parent‑child relationships between cells in the JSON output. | Provide a sample JSON snippet for a workbook where A1=5, B1=A1*2, and C1=B1+3 using the given JsonSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace FormulaTreeExport
{
    // Loads an existing workbook, configures JsonSaveOptions (ExportNestedStructure, AlwaysExportAsJsonObject, ExportAsString, ExportEmptyCells, HasHeaderRow, Indent) and saves the file as JSON. The output contains each cell’s address, value, and formula, forming a parent‑child representation of the workbook’s calculation graph.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook or load an existing one.
            // -----------------------------------------------------------------
            // Example: load an existing workbook from disk.
            // Replace "input.xlsx" with the path to your workbook.
            Workbook workbook = new Workbook("input.xlsx");

            // -----------------------------------------------------------------
            // 2. Configure JSON save options to export a parent‑child hierarchy.
            //    ExportNestedStructure = true tells Aspose.Cells to output the
            //    data as a nested JSON structure, which includes formulas and
            //    their relationships (e.g., cells that reference other cells).
            // -----------------------------------------------------------------
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportNestedStructure = true,          // Enable parent‑child JSON
                AlwaysExportAsJsonObject = true,      // Ensure the workbook is an object even if single sheet
                ExportAsString = true,                // Export cell values as strings (preserves formulas as text)
                ExportEmptyCells = false,             // Skip empty cells to keep output concise
                HasHeaderRow = true,                  // Treat first row as header (optional)
                Indent = "  "                         // Pretty‑print with indentation
            };

            // -----------------------------------------------------------------
            // 3. Save the workbook as a JSON file. The resulting file contains
            //    a hierarchical representation of the workbook, where each
            //    cell node includes its address, value, and formula (if any).
            //    This effectively captures the formula tree (parent‑child
            //    relationships) for all calculations in the workbook.
            // -----------------------------------------------------------------
            string outputPath = "formulaTree.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook formula tree exported to: {outputPath}");
        }
    }
}
