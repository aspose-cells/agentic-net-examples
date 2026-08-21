// Title: Create an Excel Formula Dependency Graph and Export to DOT with Aspose.Cells (C#)
// Description: Loads an .xlsx workbook, enables the calculation chain, evaluates all formulas, then walks every worksheet to collect precedent cells via GetPrecedentsInCalculation. The code builds a Graphviz DOT representation where each edge links a precedent cell to its dependent formula cell, and writes the graph to a .dot file.
// Keywords: Aspose.Cells formula dependency graph | C# export Excel formulas to DOT | Graphviz Excel dependency diagram | GetPrecedentsInCalculation example | cross‑sheet formula visualization | Excel calculation chain Aspose | generate DOT file from workbook
// Common Searches: how to generate a formula dependency graph from Excel using Aspose.Cells | C# code to export Excel formula relationships to Graphviz DOT | visualize inter‑sheet formula dependencies with Aspose.Cells | create DOT file of Excel calculation chain .NET | Aspose.Cells GetPrecedentsInCalculation sample
// Developer Intent: Produce a DOT file that maps every formula cell to the cells it references across all worksheets.
// Use Cases: Document and audit complex workbook calculations with a visual diagram. | Identify circular references or hidden data flows by analyzing the directed graph. | Integrate the DOT output with Graphviz or other visualization tools for reporting.
// AI Prompts: Generate C# code using Aspose.Cells that iterates all worksheets, extracts formula precedents, and writes a Graphviz DOT file. | Show how to add custom node attributes (e.g., sheet name, cell type) to the DOT output for richer visualization. | Explain the steps to render the produced dependency_graph.dot with Graphviz to create PNG or SVG images.

using System;
using System.Collections;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsDependencyGraph
{
    // Loads an .xlsx workbook, enables the calculation chain, evaluates all formulas, then walks every worksheet to collect precedent cells via GetPrecedentsInCalculation. The code builds a Graphviz DOT representation where each edge links a precedent cell to its dependent formula cell, and writes the graph to a .dot file.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // StringBuilder to build DOT graph content
            StringBuilder dotBuilder = new StringBuilder();
            dotBuilder.AppendLine("digraph G {");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells
                        if (cell.IsFormula)
                        {
                            // Get precedents that affect this cell's calculation
                            IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();

                            if (precedentsEnum != null)
                            {
                                while (precedentsEnum.MoveNext())
                                {
                                    ReferredArea area = (ReferredArea)precedentsEnum.Current;

                                    // Iterate over each cell in the referred area
                                    for (int r = area.StartRow; r <= area.EndRow; r++)
                                    {
                                        for (int c = area.StartColumn; c <= area.EndColumn; c++)
                                        {
                                            string precedentName = $"{area.SheetName}!{CellsHelper.CellIndexToName(r, c)}";
                                            string dependentName = $"{sheet.Name}!{cell.Name}";

                                            // Add edge from precedent to dependent
                                            dotBuilder.AppendLine($"    \"{precedentName}\" -> \"{dependentName}\";");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            dotBuilder.AppendLine("}");

            // Write the DOT file
            File.WriteAllText("dependency_graph.dot", dotBuilder.ToString());

            Console.WriteLine("Dependency graph exported to dependency_graph.dot");
        }
    }
}
