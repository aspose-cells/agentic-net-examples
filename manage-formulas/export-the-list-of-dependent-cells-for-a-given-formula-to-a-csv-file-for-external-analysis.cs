// Title: Export dependent cells of a specific formula to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get dependent cells of a formula and save to CSV | How to list all cells that reference a particular formula with Aspose.Cells .NET | Export formula dependency tree to CSV using Aspose.Cells in C# | Write dependent cell addresses and their formulas to a file with Aspose.Cells | Retrieve dependent cells for cell B2 and output as CSV in .NET
// Tags: Aspose.Cells GetDependents CSV export | export formula dependents C# | list dependent cells Aspose.Cells | write cell addresses and formulas Aspose.Cells | CSV output of cell dependencies .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses a target formula cell (B2), uses GetDependents(true) to collect all cells that depend on that formula, and writes each dependent cell's address and its formula (if any) to a CSV file named dependents.csv.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "dependents.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell that contains the formula whose dependents we want to export
            Cell formulaCell = worksheet.Cells["B2"];

            // Retrieve all cells that depend on the specified formula cell (true = all dependents)
            Cell[] dependentCells = formulaCell.GetDependents(true);

            // Export the list of dependent cells to a CSV file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Write CSV header
                writer.WriteLine("DependentCell,Formula");

                // Write each dependent cell's address and its formula (if any)
                foreach (Cell depCell in dependentCells)
                {
                    string address = depCell.Name; // e.g., "C5"
                    string formula = depCell.IsFormula ? depCell.Formula : string.Empty;
                    writer.WriteLine($"{address},{formula}");
                }
            }

            Console.WriteLine($"Dependent cells have been exported to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
