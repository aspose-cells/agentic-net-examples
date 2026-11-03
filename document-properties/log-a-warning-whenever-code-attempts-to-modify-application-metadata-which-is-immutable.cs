// Title: Log a warning before modifying immutable built‑in document properties in Aspose.Cells for .NET (C#)
// Common Searches: how to display a warning when changing built‑in workbook properties with Aspose.Cells in C# | prevent accidental modification of immutable document metadata using Aspose.Cells .NET | Aspose.Cells C# example for logging warning before setting Author property | check and warn before updating built‑in document properties in an Excel file with Aspose.Cells | C# console log warning for immutable application metadata in Aspose.Cells workbook
// Tags: log warning before modifying built-in document properties Aspose.Cells | immutable workbook metadata handling C# | built-in document property update Aspose.Cells | application metadata immutability Aspose.Cells .NET | Excel workbook property modification warning C#

using System;
using Aspose.Cells;

// The sample creates a Workbook, calls a helper method that writes a console warning before attempting to change the immutable 'Author' built‑in document property, sets the property value if it exists, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Log a warning before attempting to modify built‑in document properties
            LogMetadataModificationWarning("Author");

            // Modify the built‑in document property using the correct API (set the Value)
            if (workbook.BuiltInDocumentProperties["Author"] != null)
            {
                workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to log a warning whenever metadata modification is attempted
    static void LogMetadataModificationWarning(string propertyName)
    {
        Console.WriteLine($"Warning: Attempt to modify immutable application metadata property '{propertyName}'.");
    }
}
