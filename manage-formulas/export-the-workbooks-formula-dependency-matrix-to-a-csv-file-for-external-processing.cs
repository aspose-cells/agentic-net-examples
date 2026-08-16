// Title: Export Excel formula dependency matrix to CSV using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, enables the calculation chain, forces formula evaluation, iterates all used cells to capture source‑dependent pairs via Cells.GetDependents, writes the pairs with sheet‑qualified A1 addresses to a CSV file, and saves the workbook unchanged.
// Keywords: Aspose.Cells export formula dependencies | C# GetDependents CSV | Excel dependency matrix Aspose | calculation chain Aspose.Cells | write cell relationships to CSV
// Common Searches: Aspose.Cells extract formula dependencies to CSV | How to get dependent cells with Aspose.Cells C# | Export Excel cell dependency matrix using Aspose | Enable calculation chain for dependency analysis Aspose.Cells
// Developer Intent: Retrieve every formula's source‑to‑dependent relationship from a workbook and output the data as a CSV file.
// Use Cases: Generate a dependency report for auditing complex spreadsheets before modifications. | Feed the CSV into graph‑analysis tools to visualize formula interconnections. | Create an impact‑analysis matrix to identify cells that require recalculation after changes.
// AI Prompts: Write C# code that uses Aspose.Cells to export a formula dependency matrix to CSV, including handling for missing input files. | Show how to extend the sample to include hidden rows and columns and add a timestamp column to the CSV output. | Suggest performance improvements for extracting dependencies from very large workbooks with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependencyExport
{
    // Loads an Excel workbook, enables the calculation chain, forces formula evaluation, iterates all used cells to capture source‑dependent pairs via Cells.GetDependents, writes the pairs with sheet‑qualified A1 addresses to a CSV file, and saves the workbook unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";
                const string csvPath = "dependency_matrix.csv";

                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Enable calculation chain for dependency analysis
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;

                // Ensure all formulas are calculated
                workbook.CalculateFormula();

                // List to store dependency pairs (source, dependent)
                List<(string Source, string Dependent)> dependencies = new List<(string, string)>();

                // Iterate through worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through used rows and columns
                    for (int row = 0; row <= cells.MaxDataRow; row++)
                    {
                        for (int col = 0; col <= cells.MaxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Skip empty cells
                            if (cell.Type == CellValueType.IsNull) continue;

                            // Get dependent cells
                            Cell[] dependents = cells.GetDependents(true, row, col);
                            if (dependents == null) continue;

                            foreach (Cell dependentCell in dependents)
                            {
                                // Record dependency using A1 notation with sheet name
                                string sourceAddress = $"{sheet.Name}!{cell.Name}";
                                string dependentAddress = $"{sheet.Name}!{dependentCell.Name}";
                                dependencies.Add((sourceAddress, dependentAddress));
                            }
                        }
                    }
                }

                // Export dependencies to CSV
                using (StreamWriter writer = new StreamWriter(csvPath))
                {
                    writer.WriteLine("SourceCell,DependentCell");
                    foreach (var pair in dependencies)
                    {
                        writer.WriteLine($"{pair.Source},{pair.Dependent}");
                    }
                }

                // Save the workbook (unchanged) to demonstrate lifecycle usage
                workbook.Save(outputPath);
                Console.WriteLine("Dependency extraction completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
