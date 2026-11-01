// Title: Check that all named ranges in an Excel workbook use A1 reference style with Aspose.Cells for .NET
// Common Searches: how to ensure named ranges use A1 notation in Aspose.Cells C# | C# code to detect R1C1 references in Excel named ranges using Aspose.Cells | list named ranges and validate their address format with Aspose.Cells for .NET | validate Excel named range addresses are not R1C1 style in a .NET application | Aspose.Cells verify A1 reference style for workbook names
// Tags: A1 reference style check for named ranges Aspose.Cells | R1C1 pattern detection in Excel names C# | enumerate workbook defined names Aspose.Cells | regex based Excel address validation .NET | named range address verification Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, iterates through all defined names, uses a regular expression to identify R1C1‑style references, flags each named range as valid when its RefersTo string does not match the pattern (i.e., it follows A1 notation), outputs the results, and saves the workbook.
class NamedRangeValidator
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook (no password)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Regex to detect R1C1 style references (e.g., R1C1, R[1]C[1])
        Regex r1c1Pattern = new Regex(@"R(\[?-?\d+\]?|)\s*C(\[?-?\d+\]?|)", RegexOptions.IgnoreCase);

        // Iterate through all named ranges in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            string refersTo = namedRange.RefersTo; // Example: Sheet1!$A$1:$B$2

            // Determine if the address conforms to A1 style (i.e., does NOT match R1C1 pattern)
            bool isA1Style = !r1c1Pattern.IsMatch(refersTo);

            Console.WriteLine($"Named range '{namedRange.Text}' refers to '{refersTo}' -> {(isA1Style ? "Valid A1" : "Invalid (R1C1)")}.");
        }

        try
        {
            // Save the workbook (optional, if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
