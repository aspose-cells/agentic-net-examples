// Title: Retrieve the row address of a range after using the EntireRow property with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get address of rows returned by EntireRow | How to read RefersTo string after applying EntireRow to a range in Aspose.Cells | C# obtain row numbers from a selected Excel range using Aspose.Cells | Get intersecting whole rows of a range with Aspose.Cells .NET | Extract row range (e.g., 2:4) from a range object in Aspose.Cells
// Tags: EntireRow RefersTo extraction with Aspose.Cells | C# row index extraction from Excel range | Aspose.Cells whole‑row range retrieval | Excel row address generation using Aspose.Cells | Convert range to row address Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a B2:D4 range, uses the EntireRow property to get the intersecting whole rows, reads the RefersTo address (e.g., "2:4"), and prints the row address.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Create a sample range (e.g., B2:D4)
            Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange("B2:D4");

            // Get the entire rows that intersect the range
            Aspose.Cells.Range entireRowRange = range.EntireRow;

            // Retrieve the address of the entire‑row range (e.g., "2:4")
            string address = entireRowRange.RefersTo;

            // Display the address
            Console.WriteLine(address);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
