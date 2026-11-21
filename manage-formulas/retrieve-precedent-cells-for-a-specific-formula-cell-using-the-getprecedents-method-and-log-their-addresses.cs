// Title: Retrieve and list precedent cells for a formula cell using Aspose.Cells GetPrecedents in C#
// Common Searches: C# Aspose.Cells get cells referenced by a formula | list precedent cell addresses for a given cell using Aspose.Cells | how to use GetPrecedents to find formula dependencies in Excel with C# | iterate over ReferredArea collection to output precedent cells Aspose.Cells
// Tags: Aspose.Cells GetPrecedents method | process ReferredArea collection Aspose.Cells | CellsHelper index-to-name conversion C# | extract Excel formula dependencies Aspose.Cells | enumerate precedent cells C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, selects a formula cell (e.g., C5), calls GetPrecedents to obtain all referenced areas, iterates through each ReferredArea, converts row/column indices to A1-style addresses with CellsHelper, and writes each precedent cell address to the console while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the address of the formula cell whose precedents we want
            string formulaCellAddress = "C5";

            // Get the cell object
            Cell formulaCell = sheet.Cells[formulaCellAddress];

            // Retrieve the precedent cell areas for the formula cell
            ReferredAreaCollection precedentAreas = formulaCell.GetPrecedents();

            // Log the addresses of all precedent cells
            Console.WriteLine($"Precedent cells for {formulaCellAddress}:");
            foreach (ReferredArea area in precedentAreas)
            {
                // A ReferredArea may cover multiple rows/columns; iterate through each cell in the area
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        string address = CellsHelper.CellIndexToName(row, col);
                        Console.WriteLine(address);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
