// Title: Validate that named range addresses follow A1 notation using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells verify named range RefersTo uses A1 reference style | Check Excel named range address format with Aspose.Cells .NET | Validate A1 notation of named ranges in a workbook using Aspose.Cells C# | Regex for A1 style address validation of named ranges in Aspose.Cells | Iterate workbook named ranges and ensure A1 address compliance in C#
// Tags: named range address validation Aspose.Cells | A1 reference style check .NET | iterate workbook names Aspose.Cells | regex A1 notation C# | Excel named range verification Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, iterates through all defined names, extracts each name's RefersTo formula, strips the leading '=', isolates the cell address, validates the address against a regular expression that matches A1‑style references, prints whether each address is valid, and finally saves the workbook.
class NamedRangeValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Regex that matches A1 style references (single cell or range)
        Regex a1Regex = new Regex(@"^\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$", RegexOptions.IgnoreCase);

        // Iterate through all named ranges in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // RefersTo contains the formula, e.g., "=Sheet1!$A$1"
            string refersTo = namedRange.RefersTo ?? string.Empty;

            // Remove leading '=' if present
            if (refersTo.StartsWith("="))
                refersTo = refersTo.Substring(1);

            // Separate sheet name and address part
            int exclPos = refersTo.LastIndexOf('!');
            string address = exclPos >= 0 ? refersTo.Substring(exclPos + 1) : refersTo;

            // Validate the address against the A1 regex
            bool isA1 = a1Regex.IsMatch(address);

            // Use the RefersTo string as an identifier (Name property may not be available in some versions)
            Console.WriteLine($"Named range '{refersTo}' address '{address}' is {(isA1 ? "valid A1" : "invalid")}.");
        }

        try
        {
            // Save the workbook (no modifications made, just to follow lifecycle rules)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
