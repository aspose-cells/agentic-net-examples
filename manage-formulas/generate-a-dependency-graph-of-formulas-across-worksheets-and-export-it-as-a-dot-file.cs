// Title: Generate a cross‑worksheet formula dependency graph and export it to a GraphViz DOT file using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells create formula dependency graph across multiple sheets | export Excel formula relationships to GraphViz DOT using Aspose.Cells | how to extract cell references from formulas in a workbook with Aspose.Cells | build directed graph of Excel formulas for visualization in C# | Aspose.Cells generate DOT file for formula dependency analysis
// Tags: Aspose.Cells formula dependency extraction | cross‑worksheet cell reference graph | export formula graph to DOT format | C# GraphViz DOT generation from Excel | directed graph of Excel formulas

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace FormulaDependencyGraph
{
    // The example loads an Excel workbook with Aspose.Cells, iterates all worksheets, identifies cells containing formulas, parses each formula with a regular expression to capture referenced cells (including sheet names), builds a set of nodes and directed edges representing source‑to‑target dependencies, and writes the resulting graph in GraphViz DOT syntax to a specified output file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputPath = "input.xlsx";
            // Output DOT file path
            string outputPath = "dependency_graph.dot";

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(inputPath);

            // Containers for graph nodes and edges
            HashSet<string> nodes = new HashSet<string>();
            List<(string From, string To)> edges = new List<(string, string)>();

            // Regular expression to capture cell references (with optional sheet name)
            Regex refRegex = new Regex(@"(?:(?<sheet>[^!]+)!)?(?<cell>\$?[A-Z]+\$?\d+)", RegexOptions.Compiled);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (string.IsNullOrEmpty(cell.Formula))
                        continue; // Skip non‑formula cells

                    // Full address of the formula cell (including sheet name)
                    string target = $"{sheet.Name}!{cell.Name}";
                    nodes.Add(target);

                    // Find all cell references inside the formula
                    foreach (Match match in refRegex.Matches(cell.Formula))
                    {
                        string refSheet = match.Groups["sheet"].Success ? match.Groups["sheet"].Value : sheet.Name;
                        string refCell = match.Groups["cell"].Value;
                        string source = $"{refSheet}!{refCell}";

                        nodes.Add(source);
                        edges.Add((source, target));
                    }
                }
            }

            // Write the graph in DOT format
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                writer.WriteLine("digraph FormulaDependencyGraph {");
                writer.WriteLine("    rankdir=LR;"); // optional layout direction

                // Declare nodes (optional, DOT allows implicit node creation)
                foreach (string node in nodes)
                {
                    // Escape double quotes in names
                    string safeNode = node.Replace("\"", "\\\"");
                    writer.WriteLine($"    \"{safeNode}\";");
                }

                // Declare edges
                foreach (var edge in edges)
                {
                    string from = edge.From.Replace("\"", "\\\"");
                    string to = edge.To.Replace("\"", "\\\"");
                    writer.WriteLine($"    \"{from}\" -> \"{to}\";");
                }

                writer.WriteLine("}");
            }

            Console.WriteLine($"Dependency graph exported to '{outputPath}'.");
        }
    }
}
