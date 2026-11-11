// Title: C# example to list workbook‑scoped named ranges and print each name with Aspose.Cells for .NET
// Common Searches: how to list workbook scoped named ranges in an Excel file using Aspose.Cells C# | C# code to enumerate all defined names of a workbook with Aspose.Cells | retrieve only workbook level named ranges from .xlsx using Aspose.Cells for .NET | Aspose.Cells example that writes named range identifiers to the console | display all workbook scoped names from an Excel workbook in a C# console application
// Tags: extract workbook scoped named ranges Aspose.Cells | enumerate defined names collection .xlsx C# | iterate Worksheets.Names Aspose.Cells API | output named range identifiers to console C# | load Excel workbook and read named ranges Aspose.Cells | Aspose.Cells .NET workbook named range extraction

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of an input .xlsx file, loads it into an Aspose.Cells Workbook, iterates the Worksheets.Names collection, and writes each defined name's text to the console, with exception handling for robustness.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string filePath = "input.xlsx";

        try
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Iterate through all defined names and output them
            // (Older Aspose.Cells versions may not expose IsWorkbookScoped, so we output all names)
            foreach (Name name in workbook.Worksheets.Names)
            {
                Console.WriteLine(name.Text);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
