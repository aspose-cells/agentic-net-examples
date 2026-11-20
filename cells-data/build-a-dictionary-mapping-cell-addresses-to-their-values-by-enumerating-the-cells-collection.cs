// Title: Create a C# dictionary of Excel cell addresses and their values by iterating the Cells collection with Aspose.Cells
// Common Searches: how to build a dictionary of cell address and value using Aspose.Cells in C# | enumerate all cells in a worksheet and get their names with Aspose.Cells .NET | Aspose.Cells map cell addresses to objects in a Dictionary | retrieve cell name and value pairs from an Excel file using Aspose.Cells | C# iterate worksheet.Cells collection to collect address‑value pairs
// Tags: iterate worksheet.Cells Aspose.Cells | cell address to value dictionary C# | load password‑protected workbook Aspose.Cells | extract cell names and values Aspose.Cells | map Excel cell addresses to objects .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook (including optional password handling), accesses the first worksheet, iterates through its Cells collection, and stores each cell's address (Name) together with its Value in a Dictionary<string, object>. The resulting address‑value pairs are then printed to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        Workbook workbook = null;

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Attempt to load the workbook without a password
            workbook = new Workbook(inputPath);
        }
        catch (CellsException)
        {
            // If loading fails, assume the workbook may be password‑protected
            Console.WriteLine("The workbook may be password protected. Attempting to load with a password.");

            // TODO: Replace with the actual password
            const string password = "yourPassword";

            try
            {
                var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook with password: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        try
        {
            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Map cell addresses to their values
            var addressToValue = new Dictionary<string, object>();

            foreach (Cell cell in worksheet.Cells)
            {
                addressToValue[cell.Name] = cell.Value;
            }

            // Output the collected address‑value pairs
            Console.WriteLine("Cell address‑value pairs:");
            foreach (var kvp in addressToValue)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
