using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependencyGraph
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data with formulas across worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Sample formulas
            sheet1.Cells["A1"].PutValue(10);                     // source value
            sheet1.Cells["B1"].Formula = "A1*2";                // depends on A1
            sheet1.Cells["C1"].Formula = "B1+5";                // depends on B1
            sheet2.Cells["A1"].Formula = "Sheet1!C1*3";        // cross‑sheet dependency
            sheet2.Cells["B1"].Formula = "A1+Sheet1!A1";       // cross‑sheet dependency

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Containers for graph nodes and edges
            HashSet<string> nodes = new HashSet<string>();
            HashSet<(string from, string to)> edges = new HashSet<(string, string)>();

            // Iterate through all worksheets and cells
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                // Use the maximum used row/column to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null || !cell.IsFormula) continue;

                        string cellFullName = $"{ws.Name}!{cell.Name}";
                        nodes.Add(cellFullName);

                        // Get dependents whose calculated result depends on this cell
                        IEnumerator dependents = cell.GetDependentsInCalculation(true);
                        if (dependents == null) continue;

                        while (dependents.MoveNext())
                        {
                            if (dependents.Current is Cell depCell)
                            {
                                string depFullName = $"{depCell.Worksheet.Name}!{depCell.Name}";
                                nodes.Add(depFullName);
                                edges.Add((cellFullName, depFullName));
                            }
                        }
                    }
                }
            }

            // Export the graph to DOT format
            string dotPath = "dependency_graph.dot";
            using (StreamWriter writer = new StreamWriter(dotPath))
            {
                writer.WriteLine("digraph DependencyGraph {");
                writer.WriteLine("    rankdir=LR;"); // optional layout direction

                // Declare nodes (optional, DOT allows implicit node creation)
                foreach (string node in nodes)
                {
                    writer.WriteLine($"    \"{node}\";");
                }

                // Declare edges
                foreach (var edge in edges)
                {
                    writer.WriteLine($"    \"{edge.from}\" -> \"{edge.to}\";");
                }

                writer.WriteLine("}");
            }

            // Save the workbook (demonstration of lifecycle rule)
            workbook.Save("SampleWorkbook.xlsx", SaveFormat.Xlsx);

            Console.WriteLine($"Dependency graph exported to {dotPath}");
        }
    }
}