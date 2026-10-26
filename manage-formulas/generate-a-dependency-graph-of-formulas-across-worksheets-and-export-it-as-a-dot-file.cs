// Title: Create a formula dependency graph across Excel worksheets and export it as a Graphviz DOT file using Aspose.Cells for .NET
// Common Searches: how to build a formula dependency graph from an Excel workbook with Aspose.Cells in C# | export Excel cell dependency relationships to a Graphviz DOT file using .NET | C# code to list precedent cells for formulas across multiple worksheets with Aspose.Cells | generate Graphviz DOT representation of Excel formula links in a .NET application | extract cross‑sheet formula references from .xlsx using Aspose.Cells API
// Tags: Aspose.Cells extract formula precedents | C# generate Graphviz DOT from Excel formulas | cross‑worksheet formula dependency mapping | export Excel formula graph to DOT format | regex cell reference parsing Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads or creates an Excel workbook, scans every worksheet for formula cells, uses a regular expression to capture referenced cells (including sheet names), builds a dictionary of dependent‑to‑precedent relationships, constructs a left‑to‑right Graphviz DOT representation with proper identifier escaping, and writes the dependency graph to a .dot file.
class FormulaDependencyGraphGenerator
{
    static void Main(string[] args)
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "dependency_graph.dot";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a sample workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].Formula = "=B1+C1";
                ws.Cells["B1"].PutValue(5);
                ws.Cells["C1"].PutValue(10);
                workbook.Save(inputPath);
                Console.WriteLine($"Sample workbook created at '{inputPath}'.");
            }
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        // Dictionary to store edges: key = dependent cell, value = set of precedent cells
        var dependencies = new Dictionary<string, HashSet<string>>();

        // Regex to capture cell references, with optional sheet name
        Regex refRegex = new Regex(
            @"(?:(?:'(?<sheet>[^']+)'|(?<sheet>[^'!]+))!)?(?<cell>\$?[A-Z]+\$?\d+)",
            RegexOptions.Compiled);

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue;

                    string dependent = $"{sheet.Name}!{cell.Name}";

                    if (!dependencies.ContainsKey(dependent))
                        dependencies[dependent] = new HashSet<string>();

                    string formula = cell.Formula;

                    foreach (Match match in refRegex.Matches(formula))
                    {
                        string sheetName = match.Groups["sheet"].Success
                            ? match.Groups["sheet"].Value
                            : sheet.Name;

                        string cellRef = match.Groups["cell"].Value.Replace("$", "");

                        string precedent = $"{sheetName}!{cellRef}";
                        dependencies[dependent].Add(precedent);
                    }
                }
            }

            // Build DOT representation
            StringBuilder dotBuilder = new StringBuilder();
            dotBuilder.AppendLine("digraph FormulaDependencyGraph {");
            dotBuilder.AppendLine("    rankdir=LR;");

            foreach (var kvp in dependencies)
            {
                string fromNode = kvp.Key;
                foreach (string toNode in kvp.Value)
                {
                    dotBuilder.AppendLine($"    \"{EscapeDotId(fromNode)}\" -> \"{EscapeDotId(toNode)}\";");
                }
            }

            dotBuilder.AppendLine("}");

            // Write DOT file
            File.WriteAllText(outputPath, dotBuilder.ToString());
            Console.WriteLine($"Dependency graph exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }

    // Escape double quotes and backslashes for DOT identifiers
    private static string EscapeDotId(string id)
    {
        return id.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }
}
