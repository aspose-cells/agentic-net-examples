// Title: Create an Excel Formula Dependency Graph and Export to Graphviz DOT with Aspose.Cells (C#)
// Description: This C# example builds a directed graph of formula precedents across all worksheets in an Aspose.Cells workbook. It enables the calculation chain, walks every formula cell, extracts its precedents via GetPrecedentsInCalculation, assembles (precedent → dependent) edges, and writes the result to a Graphviz DOT file. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# | Excel formula dependency | Graphviz DOT export | GetPrecedentsInCalculation | calculation chain | inter‑sheet references | dependency graph | visualize Excel formulas | dot file generation
// Common Searches: Aspose.Cells generate formula dependency graph | Export Excel cell dependencies to DOT format C# | Get precedents of formula cells across worksheets Aspose.Cells | Create Graphviz diagram from Excel formulas .NET | How to visualize Excel calculation chain with Aspose.Cells
// Developer Intent: Produce a DOT file that visualizes precedence relationships between all formula cells in an Aspose.Cells workbook, covering multiple worksheets.
// Use Cases: Identify and troubleshoot circular references in large workbooks. | Perform impact analysis before modifying inter‑sheet formulas. | Generate documentation diagrams for complex calculation flows. | Integrate the DOT output with Graphviz to create visual reports.
// AI Prompts: Write a reusable method that takes an Aspose.Cells Workbook and returns a list of (from, to) dependency tuples for every formula cell. | Enhance the DOT exporter to group cells by worksheet using subgraph clusters for clearer visual separation. | Add error handling that logs cells with invalid formulas or missing precedents while continuing graph generation.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependencyGraph
{
    // This C# example builds a directed graph of formula precedents across all worksheets in an Aspose.Cells workbook. It enables the calculation chain, walks every formula cell, extracts its precedents via GetPrecedentsInCalculation, assembles (precedent → dependent) edges, and writes the result to a Graphviz DOT file. The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data: create two worksheets with formulas
            // -------------------------------------------------
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Sheet1";
            ws1.Cells["A1"].PutValue(10);
            ws1.Cells["A2"].Formula = "=A1*2";
            ws1.Cells["B1"].Formula = "=A2+Sheet2!C1";

            Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
            ws2.Cells["C1"].PutValue(5);
            ws2.Cells["D1"].Formula = "=Sheet1!A2+5";

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Build dependency edges: (precedent) -> (dependent)
            var edges = new List<(string from, string to)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    // Process only formula cells
                    if (!cell.IsFormula) continue;

                    // Get precedents that affect this cell during calculation
                    IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();
                    if (precedentsEnum == null) continue;

                    while (precedentsEnum.MoveNext())
                    {
                        if (precedentsEnum.Current is ReferredArea area)
                        {
                            // Determine the sheet name of the precedent area
                            string precedentSheet = area.SheetName ?? sheet.Name;

                            // Iterate over all cells in the area (single cell or range)
                            for (int r = area.StartRow; r <= area.EndRow; r++)
                            {
                                for (int c = area.StartColumn; c <= area.EndColumn; c++)
                                {
                                    string precedentCellName = CellsHelper.CellIndexToName(r, c);
                                    string from = $"{precedentSheet}!{precedentCellName}";
                                    string to = $"{sheet.Name}!{cell.Name}";
                                    edges.Add((from, to));
                                }
                            }
                        }
                    }
                }
            }

            // Export the graph to DOT format
            string dotPath = "FormulaDependencyGraph.dot";
            using (StreamWriter writer = new StreamWriter(dotPath))
            {
                writer.WriteLine("digraph FormulaDependencyGraph {");
                writer.WriteLine("    rankdir=LR;"); // optional layout direction

                // Write each edge
                foreach (var edge in edges)
                {
                    // Escape node names that may contain special characters
                    string fromEsc = $"\"{edge.from}\"";
                    string toEsc = $"\"{edge.to}\"";
                    writer.WriteLine($"    {fromEsc} -> {toEsc};");
                }

                writer.WriteLine("}");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("DependencyDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine($"Dependency graph exported to '{dotPath}'.");
            Console.WriteLine("Workbook saved as 'DependencyDemo.xlsx'.");
        }
    }
}
