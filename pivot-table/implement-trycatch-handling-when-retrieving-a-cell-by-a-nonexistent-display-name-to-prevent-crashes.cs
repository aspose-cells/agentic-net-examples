// Title: How to safely retrieve a cell by display name in Aspose.Cells for .NET using try‑catch to handle missing names
// Common Searches: C# Aspose.Cells catch CellsException when named cell does not exist | prevent crash when accessing non‑existent display name in Aspose.Cells workbook | example of try‑catch around worksheet.Cells["Name"] in Aspose.Cells | handle missing named range lookup in Aspose.Cells C# | Aspose.Cells retrieve cell by name with exception handling
// Tags: exception handling CellsException Aspose.Cells | named range lookup error handling C# | retrieve cell by display name Aspose.Cells | prevent workbook crash missing cell name | worksheet.Cells indexer exception handling

using System;
using Aspose.Cells;

// The example creates a new Workbook, attempts to access a cell via a display name that may not exist, wraps the lookup in a try‑catch block that catches CellsException, logs an informative message when the name is missing, and then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: create a named range (optional)
        // worksheet.Cells.CreateRange("ExistingName", "A1:B2");

        // Try to retrieve a cell by a display name that may not exist
        try
        {
            // Replace "NonExistentName" with the display name you want to access
            Cell cell = worksheet.Cells["NonExistentName"];
            Console.WriteLine($"Cell address: {cell.Name}");
        }
        catch (CellsException ex)
        {
            // Handle the case where the display name does not exist
            Console.WriteLine($"Display name not found: {ex.Message}");
        }

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}
