// Title: Export Formula Dependent Cells to CSV Using Aspose.Cells for .NET
// Description: This C# sample builds a workbook, inserts formulas, runs calculation, obtains the cells that directly depend on a chosen cell via GetDependents, and writes their addresses to a UTF‑8 CSV file while also saving the workbook to demonstrate the lifecycle rule.
// Keywords: Aspose.Cells GetDependents | C# dependent cells CSV | formula dependency extraction .NET | export cell addresses | Aspose.Cells workbook lifecycle
// Common Searches: Aspose.Cells get cells that depend on a formula | How to export dependent cell list to CSV in C# | Retrieve direct dependents of a cell with Aspose.Cells | Save formula dependency report as CSV using Aspose.Cells | C# Aspose.Cells GetDependents example
// Developer Intent: Find the cells that recalculate when a specific source cell changes and store their references in a CSV document.
// Use Cases: Impact analysis reports for key input cells | Audit‑ready documentation of spreadsheet logic | Automated generation of dependency lists for testing | Dynamic dashboards that highlight affected cells | Feeding dependency data into external reporting pipelines
// AI Prompts: Generate C# code that extracts both direct and indirect dependents of cell C3 and outputs them as a JSON array using Aspose.Cells. | Provide a reusable method that takes any cell address and returns a CSV string of its dependent cells, including the dependent formulas. | Explain how to modify the sample to include the worksheet name and row/column indices in the CSV output.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependentsExport
{
    // This C# sample builds a workbook, inserts formulas, runs calculation, obtains the cells that directly depend on a chosen cell via GetDependents, and writes their addresses to a UTF‑8 CSV file while also saving the workbook to demonstrate the lifecycle rule.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data and formulas
            cells["A1"].PutValue(10);
            cells["B1"].Formula = "=A1*2";
            cells["C1"].Formula = "=A1+B1";
            cells["D1"].Formula = "=B1*3";
            cells["F4"].Formula = "=A1*5";

            // Calculate all formulas so that dependents are recognized
            workbook.CalculateFormula();

            // Get direct dependents of cell A1 (row 0, column 0)
            Cell[] directDependents = cells.GetDependents(false, 0, 0);

            // Export dependents to a CSV file
            string csvPath = "A1_Dependents.csv";
            using (StreamWriter writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8))
            {
                // Write CSV header
                writer.WriteLine("DependentCellName");

                // Write each dependent cell name
                foreach (Cell dep in directDependents)
                {
                    writer.WriteLine(dep.Name);
                }
            }

            // Optionally save the workbook (demonstrating the lifecycle rule)
            workbook.Save("DependentsDemo.xlsx");

            Console.WriteLine($"Direct dependents of A1 have been exported to '{csvPath}'.");
        }
    }
}
