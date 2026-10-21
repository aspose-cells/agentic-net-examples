// Title: How to retrieve a named range and loop through its cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get range by name and read each cell | C# iterate over cells in a named range using Aspose.Cells | example code for retrieving named range SalesData with Aspose.Cells | how to loop through cells of a named range in an Excel workbook with Aspose.Cells for .NET | Aspose.Cells GetRangeByName returns null handling in C#
// Tags: Aspose.Cells GetRangeByName C# | named range iteration Aspose.Cells | read cell values from named range .NET | Excel workbook named range handling Aspose.Cells | enumerate cells in Excel range C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file with Aspose.Cells, obtains the range defined by the named range "SalesData" using GetRangeByName, checks for its existence, then iterates over each Cell in the range, printing the cell address and its value. It includes basic file‑existence validation, exception handling, and an optional save step.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range associated with the named range "SalesData"
            Aspose.Cells.Range salesRange = workbook.Worksheets.GetRangeByName("SalesData");
            if (salesRange == null)
            {
                Console.WriteLine("Named range 'SalesData' not found.");
                return;
            }

            // Iterate through each cell in the range and output its address and value
            foreach (Cell cell in salesRange)
            {
                Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
