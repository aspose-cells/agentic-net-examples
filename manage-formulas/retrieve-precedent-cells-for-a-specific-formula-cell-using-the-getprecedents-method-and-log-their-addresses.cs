// Title: How to retrieve and log precedent cell addresses for a specific formula cell using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, finds the cells a formula references, and prints each referenced cell's address. | Create a C# helper that receives a worksheet and a cell reference, returns all cells referenced by that formula using Aspose.Cells, and logs their addresses. | Modify the sample to write the referenced cell addresses to a CSV file instead of the console output.
// Common Searches: Aspose.Cells C# get all cells referenced by a formula | How to list precedent cells for a specific cell using Aspose.Cells .NET | C# example for retrieving formula precedents in Aspose.Cells | Export precedent cell addresses from Excel using Aspose.Cells in C# | Iterate ReferredAreaCollection to read referenced cells in Aspose.Cells
// Tags: GetPrecedents method Aspose.Cells | enumerate dependent cells C# | Aspose.Cells referenced area enumeration | display dependent cell addresses in console | write cell list to CSV Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, selects a target formula cell, uses Aspose.Cells' GetPrecedents method to obtain all referenced areas, iterates through each ReferredArea to access individual precedent cells, and writes each cell's address to the console while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Specify the formula cell whose precedents you want to retrieve
                Cell formulaCell = worksheet.Cells["C5"]; // change to your target cell

                // Get all precedent areas for the specified formula cell
                ReferredAreaCollection precedentAreas = formulaCell.GetPrecedents();

                // Iterate through each area and each cell within the area
                foreach (ReferredArea area in precedentAreas)
                {
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            Cell precedentCell = worksheet.Cells[row, col];
                            // Log the address of each precedent cell
                            Console.WriteLine($"Precedent cell address: {precedentCell.Name}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
