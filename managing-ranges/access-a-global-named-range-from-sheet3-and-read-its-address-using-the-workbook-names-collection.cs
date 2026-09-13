// Title: Get the sheet-qualified address of a global named range on Sheet3 using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, accesses a global named range, and prints its full address including the worksheet name. | Demonstrate how to obtain a Range object from workbook.Worksheets.Names and retrieve the address string in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# read full address of a global named range | How to obtain sheet name and cell range from a named range using Aspose.Cells | C# Aspose.Cells get address of named range defined on Sheet3 | Retrieve address of workbook-level named range with Aspose.Cells .NET
// Tags: Aspose.Cells get named range address | C# workbook global named range retrieval | Aspose.Cells Worksheets.Names GetRange usage | read sheet-qualified range address Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, accesses the global named range 'MyGlobalRange' via the workbook's Worksheets.Names collection, obtains its Range object, reads the sheet-qualified address (e.g., Sheet3!A1:B2), and prints it, handling missing files or undefined names gracefully.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(filePath);

            // Access the global named range (defined on any worksheet).
            Name globalName = workbook.Worksheets.Names["MyGlobalRange"];
            if (globalName == null)
            {
                Console.WriteLine("Named range 'MyGlobalRange' not found.");
                return;
            }

            // Retrieve the Range object that the name refers to.
            Aspose.Cells.Range range = globalName.GetRange();

            // Get the address of the range (e.g., Sheet3!A1:B2).
            string address = range.Address;

            // Output the address.
            Console.WriteLine($"Address of global named range 'MyGlobalRange': {address}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
