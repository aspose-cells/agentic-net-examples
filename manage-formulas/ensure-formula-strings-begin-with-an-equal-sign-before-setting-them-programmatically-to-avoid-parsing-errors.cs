// Title: Add a leading '=' to Excel formula strings before assigning them to cells with Aspose.Cells in C# to prevent parsing errors
// AI Prompts: Write a C# utility method that checks a formula string and prepends an '=' if it is missing, then use it when setting Worksheet.Cells[].Formula with Aspose.Cells. | Refactor the example to automatically sanitize any formula input before assigning it to a cell, ensuring the workbook saves without formula parsing exceptions.
// Common Searches: Aspose.Cells C# how to add missing equal sign to formula before setting cell value | prevent formula parsing error when using Worksheet.Cells[].Formula in Aspose.Cells | C# helper function to ensure Excel formulas start with '=' for Aspose.Cells workbooks
// Tags: prepend equal sign to Excel formula Aspose.Cells | formula sanitization before setting cell C# | avoid formula parsing errors Aspose.Cells | validate Excel formula string Aspose.Cells | cell formula helper method C#

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, trimming a raw formula, ensuring it begins with '=', assigning it to cell B1, and saving as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Example formula string that may be missing the leading equal sign
        string rawFormula = "SUM(A1:A10)";

        // Ensure the formula string starts with '=' to avoid parsing errors
        string safeFormula = EnsureFormulaStartsWithEqual(rawFormula);

        // Set the formula to cell B1
        cells["B1"].Formula = safeFormula;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }

    // Helper method to prepend '=' if it's not already present
    static string EnsureFormulaStartsWithEqual(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
            return formula; // Return as is if null or empty

        // Trim leading whitespace
        string trimmed = formula.TrimStart();

        // If the trimmed formula already starts with '=', return it unchanged
        if (trimmed.StartsWith("="))
            return trimmed;

        // Otherwise, prepend '='
        return "=" + trimmed;
    }
}
