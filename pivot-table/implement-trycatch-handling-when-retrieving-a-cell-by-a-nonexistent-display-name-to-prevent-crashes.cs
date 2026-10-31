// Title: How to safely retrieve a cell by display name with try‑catch handling in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get cell by named range with exception handling | catch CellsException when a display name is missing in Aspose.Cells | prevent crash when accessing a non‑existent named cell in Aspose.Cells | use try‑catch around sheet.Cells["DisplayName"] in Aspose.Cells | handle missing named range error in Aspose.Cells workbook
// Tags: Aspose.Cells try‑catch CellsException | retrieve cell by display name Aspose.Cells | handle missing named range Aspose.Cells | prevent workbook crash Aspose.Cells C# | exception handling for cell lookup Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, attempts to access a cell using a display name that may not exist, wraps the lookup in a try‑catch block that catches CellsException, reports the missing name, and then saves the workbook without crashing.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Example: add a named range (optional)
        // sheet.Cells["B2"].PutValue(42);
        // workbook.Names.Add("ExistingName", sheet.Cells["B2"]);

        // Try to get a cell by a display name that might not exist
        try
        {
            // Replace "NonExistentName" with the display name you want to retrieve
            Cell targetCell = sheet.Cells["NonExistentName"];
            // If no exception, the cell was found – you can work with it here
            Console.WriteLine($"Cell found at address: {targetCell.Name}");
        }
        catch (CellsException ex)
        {
            // Handle the case where the display name does not exist
            Console.WriteLine($"Display name not found: {ex.Message}");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("Result.xlsx");
    }
}
