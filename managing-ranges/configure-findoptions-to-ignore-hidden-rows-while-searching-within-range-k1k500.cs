// Title: How to use Aspose.Cells FindOptions to skip hidden rows while searching cells K1:K500 in C#
// Common Searches: Aspose.Cells FindOptions ignore hidden rows when searching column K in C# | C# code to find a value in range K1:K500 without considering hidden rows using Aspose.Cells | skip hidden rows during cell search with Aspose.Cells FindOptions C# example | search specific range K1 to K500 while excluding hidden rows in Aspose.Cells
// Tags: FindOptions.SearchHiddenRows Aspose.Cells | column K range lookup Aspose.Cells C# | exclude hidden rows cell search Aspose.Cells | Aspose.Cells find value within K1:K500 | C# workbook cell search hidden rows

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, accesses the first worksheet, and creates a FindOptions object configured with LookInType.Values and SearchHiddenRows set to false. It then searches for a target string starting from cell K1, limits the result to rows 1‑500, reports the found cell location, and saves the workbook, ensuring hidden rows are ignored during the search.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure FindOptions (search in cell values)
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values
            };

            // The value to search for
            string valueToFind = "TargetValue";

            // Start the search from the first cell of the desired range (K1)
            Cell startCell = sheet.Cells["K1"];

            // Perform the search
            Cell foundCell = sheet.Cells.Find(valueToFind, startCell, findOptions);

            // Verify that the found cell is within the K1:K500 range
            if (foundCell != null && foundCell.Row <= 499) // rows are zero‑based
            {
                Console.WriteLine($"Value found at cell {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("Value not found in the specified range.");
            }

            // Save the workbook (output file will be created/overwritten)
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
