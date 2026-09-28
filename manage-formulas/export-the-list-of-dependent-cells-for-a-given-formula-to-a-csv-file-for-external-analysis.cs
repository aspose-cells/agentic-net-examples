// Title: Export dependent (precedent) cells of a specific Excel formula to a CSV file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, identifies all precedent cells of a given formula cell, and writes each cell's address and displayed value to a CSV file using Aspose.Cells. | Modify the example to accept the worksheet name, formula address, and output CSV path as command‑line arguments, then export the dependent cells accordingly. | Create a reusable C# method that returns a collection of cell addresses and values for any formula cell in an Aspose.Cells workbook, and demonstrate writing that collection to a CSV file.
// Common Searches: how to list cells referenced by a formula and save to csv using Aspose.Cells in C# | Aspose.Cells GetPrecedents export to csv example | C# extract precedent cells of A1 and write to csv file | save dependent cells of an Excel formula as csv with Aspose.Cells .NET | retrieve formula precedents and export values to csv programmatically
// Tags: export precedent cells to CSV Aspose.Cells | Aspose.Cells GetPrecedents to CSV | C# extract formula dependent cells | write Excel cell addresses and values to CSV | Aspose.Cells dependent cells extraction

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The program loads an Excel workbook, obtains the precedent cells of a specified formula (e.g., A1) on a worksheet, and writes each cell's address and displayed value to a CSV file named dependent_cells.csv using Aspose.Cells for .NET.
class ExportDependentCells
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "dependent_cells.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Specify the cell that contains the formula (e.g., "A1")
            Cell formulaCell = worksheet.Cells["A1"];

            // Get the precedents (cells referenced by the formula)
            ReferredAreaCollection precedentAreas = formulaCell.GetPrecedents();

            // Prepare CSV content
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("Address,Value"); // Header

            // Iterate through each referenced area and each cell within the area
            foreach (ReferredArea area in precedentAreas)
            {
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        string address = CellsHelper.CellIndexToName(row, col);
                        string value = cell.StringValue ?? string.Empty; // Get displayed value

                        // Escape double quotes for CSV
                        string escapedValue = value.Replace("\"", "\"\"");
                        csvBuilder.AppendLine($"{address},\"{escapedValue}\"");
                    }
                }
            }

            // Write the CSV to the output file
            File.WriteAllText(outputPath, csvBuilder.ToString());

            Console.WriteLine($"Dependent cells exported to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
