// Title: Retrieve the "SalesData" named range and iterate its cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get named range SalesData from workbook | loop through cells of a named range using Aspose.Cells in .NET | read cell addresses and values from a named range with Aspose.Cells C# | how to handle missing named range in Aspose.Cells C# code | example of using GetRangeByName and foreach cell in Aspose.Cells
// Tags: Aspose.Cells GetRangeByName C# | iterate cells in named range Aspose.Cells | read named range values .NET | Excel workbook named range processing Aspose.Cells | C# foreach cell in Aspose.Cells range

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook with Aspose.Cells, retrieves the "SalesData" named range via GetRangeByName, checks for its existence, and iterates each cell in the range, outputting the cell address and value.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range associated with the named range "SalesData"
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("SalesData");
            if (range == null)
            {
                Console.WriteLine("Named range 'SalesData' not found.");
                return;
            }

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Example processing: output cell address and value
                Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
            }

            // Save the workbook if any changes were made (optional)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
