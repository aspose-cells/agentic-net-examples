// Title: Export dependent cells of a given formula to a CSV file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export formula precedents to CSV | How to list cells referenced by a formula and save as CSV using Aspose.Cells | C# get precedent cell addresses from a workbook and write to CSV | Export dependent cell addresses of B2 formula to CSV with Aspose.Cells | Save Excel formula dependent cells to CSV file in .NET
// Tags: Aspose.Cells GetPrecedents to CSV | C# extract precedent cell addresses | write Excel cell references to CSV file | list dependent cells of a formula Aspose.Cells | save formula referenced cells as CSV .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an XLSX workbook, verifies a target cell contains a formula, retrieves all precedent cell addresses via GetPrecedents, and writes each address to a CSV file with a header row.
class ExportDependentCells
{
    static void Main()
    {
        // Paths
        string sourcePath = "input.xlsx";
        string csvPath = "dependent_cells.csv";
        string formulaCellAddress = "B2";

        try
        {
            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the cell containing the formula
            Cell formulaCell = sheet.Cells[formulaCellAddress];

            // Ensure the cell has a formula
            if (!formulaCell.IsFormula)
            {
                Console.WriteLine($"Cell {formulaCellAddress} does not contain a formula.");
                return;
            }

            // Retrieve precedent areas (cells referenced by the formula)
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents == null || precedents.Count == 0)
            {
                Console.WriteLine($"No precedent cells found for {formulaCellAddress}.");
                return;
            }

            // Write dependent cell addresses to CSV
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                writer.WriteLine("DependentCell");

                foreach (ReferredArea area in precedents)
                {
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            string address = CellsHelper.CellIndexToName(row, col);
                            writer.WriteLine(address);
                        }
                    }
                }
            }

            Console.WriteLine($"Dependent cells exported to {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
