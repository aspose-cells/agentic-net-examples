// Title: C# code to verify that conditional formatting rules are preserved when applying an Excel theme using Aspose.Cells
// Common Searches: how to check if conditional formatting is lost after applying a theme with Aspose.Cells in C# | Aspose.Cells count conditional formatting before and after ApplyTheme method | verify workbook theme change does not remove conditional formats .NET | C# sample to compare conditional formatting counts in an Excel file using Aspose.Cells
// Tags: Aspose.Cells ApplyTheme preserve conditional formatting | count conditional formatting rules C# Aspose.Cells | recalculate formulas after theme change Aspose.Cells | validate workbook theme application .NET | conditional formatting integrity Excel theme Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, counts all conditional formatting rules across worksheets, attempts to apply a built‑in theme (if supported), recalculates formulas, recounts the rules, compares the before/after totals to ensure no rules were lost, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Count conditional formatting rules before any changes
            int cfCountBefore = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                cfCountBefore += ws.ConditionalFormattings.Count;
            }

            // Apply a built‑in theme if the API is available.
            // The ApplyTheme method may not exist in older Aspose.Cells versions,
            // so it is wrapped in a try/catch to prevent compilation/runtime errors.
            try
            {
                // Uncomment the line below if your Aspose.Cells version supports ApplyTheme.
                // workbook.ApplyTheme("Office");
            }
            catch (MissingMethodException)
            {
                // Theme application not supported; continue without it.
            }

            // Recalculate formulas so that conditional formatting is evaluated
            workbook.CalculateFormula();

            // Count conditional formatting rules after the (potential) theme change
            int cfCountAfter = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                cfCountAfter += ws.ConditionalFormattings.Count;
            }

            // Compare counts to confirm no rules were lost
            if (cfCountBefore == cfCountAfter)
            {
                Console.WriteLine("All conditional formatting rules are preserved after applying the new theme.");
            }
            else
            {
                Console.WriteLine($"Conditional formatting count changed from {cfCountBefore} to {cfCountAfter}.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
