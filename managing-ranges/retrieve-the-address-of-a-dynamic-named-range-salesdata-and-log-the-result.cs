// Title: How to retrieve and log the address of a dynamic named range "SalesData" with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens an Excel workbook, fetches the dynamic named range "SalesData", obtains its address, and prints it to the console. | Show C# error‑handling patterns for when the file "input.xlsx" or the named range "SalesData" cannot be found while using Aspose.Cells. | Adapt the sample to write the resolved range address to a log file instead of the console, employing standard .NET logging facilities.
// Common Searches: Aspose.Cells C# get address of a dynamic named range from workbook | C# example for reading named range SalesData address using Aspose.Cells | How to handle missing named range when using Aspose.Cells in .NET | Log Excel named range address to a file with Aspose.Cells C#
// Tags: Aspose.Cells retrieve named range location | C# resolve named range to Range object | Excel workbook named range extraction .NET | Aspose.Cells missing named range handling | log range location using .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Loads 'input.xlsx', accesses the workbook's Names collection to locate the dynamic named range 'SalesData', resolves it to a Range object, reads its address (e.g., "A2:C15"), and outputs the address to the console while gracefully handling missing files or missing named ranges.
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

            // Retrieve the dynamic named range "SalesData"
            // Named ranges are stored in the Names collection of the worksheet collection
            Name salesDataName = workbook.Worksheets.Names["SalesData"];
            if (salesDataName == null)
            {
                Console.WriteLine("Named range 'SalesData' not found.");
                return;
            }

            // Resolve the named range to an actual Range object based on current data
            AsposeRange salesDataRange = salesDataName.GetRange();

            // Get the address of the resolved range (e.g., "A2:C15")
            string address = salesDataRange.Address;

            // Log the result
            Console.WriteLine($"SalesData address: {address}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
