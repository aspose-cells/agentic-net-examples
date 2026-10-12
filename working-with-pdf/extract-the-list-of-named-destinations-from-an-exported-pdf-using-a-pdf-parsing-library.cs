// Title: How to list all named ranges in an Excel workbook and get the first cell address using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, accesses workbook.Worksheets.Names, and prints each named range's name, RefersTo formula, and the first cell address. | Create a C# helper method that parses a RefersTo string (e.g., "=Sheet1!$A$1:$B$2") and returns the first cell reference, then use it while enumerating named ranges in an Aspose.Cells workbook.
// Common Searches: aspacells c# list named ranges and retrieve first cell address | extract first cell from named range RefersTo using Aspose.Cells | iterate workbook named collections Aspose.Cells .NET example
// Tags: Aspose.Cells enumerate named ranges | C# extract first cell from RefersTo | Aspose.Cells workbook named collections | read Excel named ranges Aspose.Cells | parse RefersTo formula C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, accesses the NameCollection via workbook.Worksheets.Names, iterates each Name object to obtain its Text (the range name) and RefersTo formula, uses a helper method to extract the first cell address from the RefersTo string, and writes these details to the console.
class Program
{
    static void Main()
    {
        string filePath = "exported.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the collection of named ranges
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Iterate through each named range and output its details
            foreach (Name range in namedRanges)
            {
                // Named range name
                string name = range.Text;

                // The reference formula (e.g., =Sheet1!$A$1:$B$2)
                string refersTo = range.RefersTo;

                // Extract the first cell address from the reference
                string firstCell = GetFirstCellAddress(refersTo);

                Console.WriteLine($"Name: {name}, RefersTo: {refersTo}, FirstCell: {firstCell}");
            }
        }
        catch (Exception ex)
        {
            // Runtime safety: capture and display any errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to extract the first cell address from a RefersTo string
    private static string GetFirstCellAddress(string refersTo)
    {
        if (string.IsNullOrEmpty(refersTo))
            return string.Empty;

        // Remove leading '=' if present
        string cleaned = refersTo.TrimStart('=');

        // Separate sheet name if present
        int exclPos = cleaned.IndexOf('!');
        if (exclPos >= 0)
            cleaned = cleaned.Substring(exclPos + 1);

        // Get the first cell before any range delimiter ':'
        int colonPos = cleaned.IndexOf(':');
        if (colonPos >= 0)
            cleaned = cleaned.Substring(0, colonPos);

        return cleaned;
    }
}
