// Title: Retrieve all cells dependent on a specific formula in an Excel workbook using Aspose.Cells for .NET and export the addresses to a CSV file
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, calls GetDependents(true) on a target cell, and writes each dependent cell's full address (Sheet!Cell) to a CSV file. | Create a .NET console application that checks the input file exists, extracts recursive dependents of cell B2 using Aspose.Cells, and saves the results with a header to dependents.csv. | Generate a C# snippet that handles exceptions, uses Aspose.Cells to find all cells that depend on a formula, and outputs the list in CSV format including sheet names.
// Common Searches: aspocells getdependents example c# recursive dependent cells | how to export list of dependent cells from Excel to CSV using Aspose.Cells | c# code to find cells that reference a formula cell with Aspose.Cells | save dependent cell addresses to CSV file in .NET Aspose.Cells
// Tags: Aspose.Cells GetDependents to CSV | export formula dependents Aspose.Cells .NET | retrieve dependent cells Excel C# | recursive cell dependency extraction Aspose.Cells | write dependent cell addresses CSV C#

using Aspose.Cells;
using System;
using System.IO;

// // Loads an Excel workbook, retrieves all cells that depend on cell B2 via GetDependents(true), and writes each dependent cell's full address (SheetName!CellName) to a CSV file.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Specify the cell that contains the formula (adjust address as needed)
            Cell formulaCell = worksheet.Cells["B2"];

            // Retrieve all cells that depend on the formula cell
            // Pass 'true' to get all dependent cells recursively
            Cell[] dependentCells = formulaCell.GetDependents(true);

            // Export the list of dependent cells to a CSV file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Write CSV header
                writer.WriteLine("DependentCell");

                // Write each dependent cell's full address (SheetName!CellName)
                foreach (Cell dep in dependentCells)
                {
                    string fullAddress = $"{dep.Worksheet.Name}!{dep.Name}";
                    writer.WriteLine(fullAddress);
                }
            }

            Console.WriteLine($"Dependent cells have been written to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
