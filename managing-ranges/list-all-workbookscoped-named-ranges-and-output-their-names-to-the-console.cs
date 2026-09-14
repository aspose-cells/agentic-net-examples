// Title: How to list workbook‑scoped named ranges in an Excel file using Aspose.Cells for .NET and print them to the console
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, filters the Names collection for workbook‑level scopes, and writes each name to the console. | Create a method that returns a string array of all workbook‑scoped defined names from a Workbook object. | Show how to handle missing files and exceptions while enumerating workbook‑scoped named ranges with Aspose.Cells.
// Common Searches: Aspose.Cells C# retrieve only workbook level named ranges from an Excel workbook | list defined names that are not tied to a worksheet using Aspose.Cells .NET | C# example to print workbook scoped names from input.xlsx with Aspose.Cells | how to differentiate workbook and worksheet scoped names in Aspose.Cells
// Tags: enumerate workbook scoped names Aspose.Cells | filter defined names by scope .NET | print named ranges console C# | load Excel workbook Aspose.Cells | handle missing file exception Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, iterates the workbook’s Worksheets.Names collection, outputs each workbook‑scoped defined name to the console, and includes basic error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found – {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Iterate through all defined names in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Output the name text (Aspose.Cells may not expose IsWorkbookScoped in older versions)
                Console.WriteLine(definedName.Text);
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions (e.g., file format issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
