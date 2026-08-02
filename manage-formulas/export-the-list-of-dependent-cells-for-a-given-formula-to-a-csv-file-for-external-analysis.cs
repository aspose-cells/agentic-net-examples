// Title: Export Formula Dependent Cells to CSV with Aspose.Cells for .NET (C#)
// Description: The sample loads an Excel workbook, triggers full formula calculation, extracts every cell that references the specified address (including other worksheets) using GetDependents(true), and writes the sheet name, address, formula and value into a CSV file; the workbook can then be saved.
// Keywords: Aspose.Cells C# export dependent cells | GetDependents Aspose.Cells | formula dependency CSV | Excel cell dependents .NET | write cell list to CSV | calculate formulas Aspose.Cells | cross‑sheet dependencies Aspose.Cells
// Common Searches: Aspose.Cells get dependent cells and export to CSV | C# retrieve cells that reference a formula | How to list downstream cells of a given address in Excel using Aspose | Export Excel formula dependency tree to CSV | GetDependents true across worksheets Aspose.Cells
// Developer Intent: Create a CSV report of all cells that reference a particular formula in an Excel workbook.
// Use Cases: Audit the impact of a critical calculation by exporting its downstream cells for manual review. | Produce a dependency ledger for financial models, capturing sheet names, addresses, formulas and values in CSV format. | Feed cell‑reference data into external analytics pipelines or visualization tools.
// AI Prompts: Generate C# code that uses Aspose.Cells to list dependents of cell B2, keep only those containing formulas, and output the result as JSON. | Explain how to modify the example for handling very large workbooks efficiently while streaming dependent rows to a CSV file. | Provide a step‑by‑step guide to add robust error handling for missing worksheets, invalid cell addresses, and duplicate entries during export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample loads an Excel workbook, triggers full formula calculation, extracts every cell that references the specified address (including other worksheets) using GetDependents(true), and writes the sheet name, address, formula and value into a CSV file; the workbook can then be saved.
    public class ExportDependentsToCsv
    {
        /// <param name="inputFilePath">Path to the source Excel workbook.</param>
        /// <param name="cellAddress">Address of the cell whose dependents are required (e.g., "A1").</param>
        /// <param name="outputCsvPath">Path where the CSV file will be created.</param>
        public static void Run(string inputFilePath, string cellAddress, string outputCsvPath)
        {
            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the workbook (load rule)
                Workbook workbook = new Workbook(inputFilePath);

                // Ensure all formulas are calculated so that dependency information is up‑to‑date
                workbook.CalculateFormula();

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Retrieve the target cell by its address
                Cell targetCell = cells[cellAddress];

                // Get all dependent cells (recursive = true, includes other worksheets)
                Cell[] dependents = targetCell.GetDependents(true);

                // Write the dependent cell information to a CSV file
                using (StreamWriter writer = new StreamWriter(outputCsvPath))
                {
                    // Header row
                    writer.WriteLine("Worksheet,CellName,Formula,Value");

                    // Iterate through each dependent cell and output its details
                    foreach (Cell dep in dependents)
                    {
                        string sheetName = dep.Worksheet.Name;
                        string name = dep.Name;
                        string formula = dep.IsFormula ? dep.Formula : string.Empty;
                        string value = dep.StringValue.Replace("\"", "\"\""); // Escape quotes

                        // CSV line (values are quoted to handle commas)
                        writer.WriteLine($"\"{sheetName}\",\"{name}\",\"{formula}\",\"{value}\"");
                    }
                }

                // Optionally, save the workbook after processing (save rule)
                workbook.Save("ProcessedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            // Expected arguments: inputFilePath cellAddress outputCsvPath
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: <inputFilePath> <cellAddress> <outputCsvPath>");
                return;
            }

            string inputFilePath = args[0];
            string cellAddress = args[1];
            string outputCsvPath = args[2];

            try
            {
                ExportDependentsToCsv.Run(inputFilePath, cellAddress, outputCsvPath);
                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }
    }
}
