// Title: C# – List All Named Ranges in an Aspose.Cells Workbook and Write Name‑Address to Debug Log
// Description: Shows how to create a workbook, add named ranges, retrieve every defined range with Workbook.Worksheets.GetNamedRanges(), and output each range’s Name and Address to the Visual Studio debug console (with optional file save).
// Keywords: Aspose.Cells | C# | named ranges | GetNamedRanges | debug log | list ranges | enumerate named ranges | Workbook.Worksheets | range address | Excel automation
// Common Searches: Aspose.Cells get all named ranges C# | list named ranges address Aspose.Cells | debug write named range name address | enumerate named ranges .NET | Aspose.Cells GetNamedRanges loop example | how to retrieve named ranges in C#
// Developer Intent: The developer wants to loop through every named range in a workbook and output its identifier and cell address for verification, reporting, or troubleshooting.
// Use Cases: Validate that required named ranges exist before running calculations. | Generate a quick documentation report of all defined ranges by logging name‑address pairs. | Debug formula errors by exposing the underlying range definitions in the debug console. | Export range metadata to another system (e.g., CSV) after enumeration.
// AI Prompts: Write C# code using Aspose.Cells to list all named ranges and save the name‑address pairs to a CSV file. | Show how to filter named ranges by a prefix (e.g., "Report_") and log only those matching ranges. | Explain how to safely handle a null return from GetNamedRanges and prevent runtime exceptions.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Shows how to create a workbook, add named ranges, retrieve every defined range with Workbook.Worksheets.GetNamedRanges(), and output each range’s Name and Address to the Visual Studio debug console (with optional file save).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample named ranges for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.CreateRange("A1:B2").Name = "SalesData";
            sheet.Cells.CreateRange("C3:D4").Name = "Expenses";

            // Retrieve all defined named ranges in the workbook
            // Use fully qualified type to avoid ambiguity with System.Range
            Aspose.Cells.Range[] namedRanges = workbook.Worksheets.GetNamedRanges();

            // Iterate through each named range and write its name and address to the debug log
            if (namedRanges != null)
            {
                foreach (Aspose.Cells.Range range in namedRanges)
                {
                    Debug.WriteLine($"Name: {range.Name}, Address: {range.Address}");
                }
            }

            // Save the workbook (optional)
            workbook.Save("NamedRangesDemo.xlsx");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Debug.WriteLine($"Error: {ex.Message}");
        }
    }
}
