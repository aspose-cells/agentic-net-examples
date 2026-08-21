// Title: C# – Retrieve the "SalesData" named range as a Range object and loop through its cells with Aspose.Cells
// Description: Load or create a workbook, get the Name object for "SalesData", convert it to an Aspose.Cells.Range, enumerate each Cell to display its address and value, and save the workbook. Includes error handling for missing named ranges.
// Keywords: Aspose.Cells named range | C# get Range object by name | iterate cells in named range | SalesData range example | read cell values Aspose.Cells | check named range existence
// Common Searches: how to retrieve a named range in Aspose.Cells .NET | loop through cells of a named range using C# | Aspose.Cells get range object from name | validate named range before reading values | save workbook after processing named range
// Developer Intent: Obtain the Range object for the "SalesData" named range and iterate each cell to read its address and value.
// Use Cases: Read and display all values in a predefined named range. | Verify a named range exists before performing calculations. | Update cell contents within a named range and persist changes.
// AI Prompts: Generate C# code with Aspose.Cells that checks for a named range, retrieves its Range object, and prints each cell's address and value. | Show how to handle a missing named range gracefully while iterating cells in Aspose.Cells. | Explain how to modify values inside a retrieved Range object and save the workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load or create a workbook, get the Name object for "SalesData", convert it to an Aspose.Cells.Range, enumerate each Cell to display its address and value, and save the workbook. Includes error handling for missing named ranges.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if the file exists; otherwise create a new workbook.
            Workbook workbook;
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Retrieve the named range "SalesData".
            Name salesName = workbook.Worksheets.Names["SalesData"];
            if (salesName == null)
            {
                Console.WriteLine("Named range 'SalesData' not found.");
                return;
            }

            // Get the Aspose.Cells.Range object that the name refers to.
            Aspose.Cells.Range salesRange = salesName.GetRange();

            // Iterate through each cell in the range and display its address and value.
            foreach (Cell cell in salesRange)
            {
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // Save the workbook (if any changes were made).
            const string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
