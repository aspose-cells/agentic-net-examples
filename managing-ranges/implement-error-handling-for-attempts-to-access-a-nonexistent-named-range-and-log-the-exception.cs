// Title: How to catch CellsException for a missing named range and log the error in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells GetRangeByName throws exception when name not found | handle missing named range error in Aspose.Cells workbook | log CellsException details for non‑existent named range in .NET | try‑catch pattern for Aspose.Cells named range retrieval | saving workbook after handling named range exception Aspose.Cells
// Tags: Aspose.Cells named range exception handling | GetRangeByName missing range error C# | log CellsException Aspose.Cells | Workbook save with error handling Aspose.Cells | C# try‑catch Aspose.Cells range retrieval

using System;
using Aspose.Cells;

// The example creates a new Workbook, attempts to retrieve a named range that does not exist using GetRangeByName, catches the resulting CellsException, logs the exception message to the console, and then saves the workbook with a separate try‑catch block to handle any save‑time errors.
class NamedRangeErrorHandling
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to access a named range that does not exist.
            // This call throws a CellsException when the name is not found.
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("NonExistentRange");
        }
        catch (CellsException ex)
        {
            // Log the exception details.
            Console.WriteLine($"Error accessing named range: {ex.Message}");
        }

        // Save the workbook (lifecycle rule: save)
        try
        {
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
