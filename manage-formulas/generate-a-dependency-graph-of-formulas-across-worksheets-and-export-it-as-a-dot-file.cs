// Title: Generate a cross‑worksheet formula dependency graph and export it as a GraphViz DOT file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, scans every worksheet for formula cells, extracts the referenced cells, and builds a GraphViz DOT representation of the dependencies. | Extend the example to include defined ranges in the analysis and output the resulting graph in Mermaid markdown instead of DOT. | Add robust error handling that logs cells with unparsable formulas while still producing a complete dependency graph file.
// Common Searches: how to visualize Excel formula relationships with Aspose.Cells in C# | C# extract cell references from formulas across multiple sheets using Aspose.Cells | export Excel cell dependency map to GraphViz format with .NET | parse cross‑sheet formulas and generate DOT file programmatically | Aspose.Cells example for building a formula connection diagram
// Tags: extract Excel formula references with Aspose.Cells C# | create DOT file for formula relationships using GraphViz | parse cross‑sheet cell addresses via regular expression C# | process defined ranges in formula graph Aspose.Cells | log unparsable formulas during dependency extraction .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet to find formula cells, uses a regular expression to identify same‑sheet and cross‑sheet cell references, constructs a GraphViz DOT representation of the dependency graph, and writes the result to a .dot file.
class FormulaDependencyGraph
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "dependency_graph.dot";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Build DOT file content
            StringBuilder dotBuilder = new StringBuilder();
            dotBuilder.AppendLine("digraph G {");

            // Track declared nodes to prevent duplicates
            HashSet<string> nodes = new HashSet<string>();

            // Regex for sheet and cell references (sheet part optional)
            Regex refRegex = new Regex(
                @"(?:(?:'(?<sheet>[^']+)')|(?<sheet>[^'!]+))?!?(?<cell>\$?[A-Za-z]+\$?\d+)",
                RegexOptions.Compiled);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula) continue;

                    // Target node identifier (sheet name + cell address)
                    string targetNode = $"{EscapeId(sheet.Name)}!{cell.Name}";
                    if (nodes.Add(targetNode))
                    {
                        dotBuilder.AppendLine($"    \"{targetNode}\";");
                    }

                    // Parse the formula text for all references (same‑sheet and cross‑sheet)
                    string formula = cell.Formula;
                    if (string.IsNullOrEmpty(formula)) continue;

                    foreach (Match match in refRegex.Matches(formula))
                    {
                        string refSheet = match.Groups["sheet"].Success && !string.IsNullOrEmpty(match.Groups["sheet"].Value)
                                          ? match.Groups["sheet"].Value
                                          : sheet.Name; // default to current sheet if not specified

                        string refCell = match.Groups["cell"].Value.Replace("$", ""); // remove absolute markers

                        string sourceNode = $"{EscapeId(refSheet)}!{refCell}";
                        if (nodes.Add(sourceNode))
                        {
                            dotBuilder.AppendLine($"    \"{sourceNode}\";");
                        }
                        dotBuilder.AppendLine($"    \"{sourceNode}\" -> \"{targetNode}\";");
                    }
                }
            }

            dotBuilder.AppendLine("}");

            // Write DOT representation to file
            try
            {
                File.WriteAllText(outputPath, dotBuilder.ToString());
                Console.WriteLine($"Dependency graph written to \"{outputPath}\".");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write output file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to escape double quotes and backslashes in DOT identifiers
    private static string EscapeId(string id)
    {
        return id.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }
}
