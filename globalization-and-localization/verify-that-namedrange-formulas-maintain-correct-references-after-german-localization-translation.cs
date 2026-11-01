// Title: Check that named range RefersTo formulas stay accurate after translating a workbook to German (de‑DE) with Aspose.Cells for .NET
// Common Searches: how to validate named range RefersTo formulas in a German localized Excel file using Aspose.Cells .NET | Aspose.Cells .NET verify named range references after setting workbook culture to de-DE | C# code to iterate named ranges and resolve RefersTo strings in a de‑DE workbook | detect invalid named range formulas after translating Excel to German with Aspose.Cells | sample program to load workbook, set CultureInfo to de-DE and check named range ranges
// Tags: named range reference validation Aspose.Cells | German locale formula parsing Aspose.Cells | create range from RefersTo C# | Aspose.Cells workbook culture configuration | resolve RefersTo string to Range object

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, applies the German (de-DE) culture for formula handling, iterates through all defined names, extracts each name's RefersTo formula, creates a Range object from the reference, verifies the range is valid, reads the first cell value as a sanity check, logs any parsing errors, and saves the workbook.
class VerifyNamedRangeFormulas
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "{InputPath}";
            string outputPath = "{OutputPath}";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            // Load the workbook that has been translated to German localization
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook uses German culture for formula parsing/display
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Access the collection of named ranges in the workbook
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Iterate through each named range and verify its formula reference
            foreach (Name namedRange in namedRanges)
            {
                // The formula that the name refers to (e.g., =Tabelle1!$A$1:$B$10)
                string refersTo = namedRange.RefersTo;

                // Output the name and its formula for inspection
                Console.WriteLine($"Name: {namedRange.Text}");
                Console.WriteLine($"  RefersTo: {refersTo}");

                // Remove leading '=' if present (Aspose stores it with '=')
                string cleanRef = refersTo.TrimStart('=');

                try
                {
                    // Create a Range object from the reference string
                    // (Works with German sheet names and cell references)
                    AsposeRange range = workbook.Worksheets[0].Cells.CreateRange(cleanRef);

                    // Verify that the range is valid
                    if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
                    {
                        // Read the value of the first cell as a sanity check
                        object firstValue = range[0, 0].Value;
                        Console.WriteLine($"  First cell value: {firstValue}");
                    }
                    else
                    {
                        Console.WriteLine("  Warning: Resolved range is empty or invalid.");
                    }
                }
                catch (Exception ex)
                {
                    // If the reference cannot be parsed, report the issue
                    Console.WriteLine($"  Error resolving reference: {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save a copy of the workbook after verification
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
