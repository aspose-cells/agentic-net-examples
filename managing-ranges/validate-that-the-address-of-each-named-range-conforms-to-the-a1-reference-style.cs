// Title: Check that each named range’s RefersTo address conforms to A1 notation using Aspose.Cells for .NET
// AI Prompts: Loop through workbook.Worksheets.Names and use CellArea.CreateCellArea to determine if each named range’s RefersTo string is a valid A1 reference. | Refactor the sample to throw a custom InvalidNamedRangeException when a RefersTo address does not meet A1 validation criteria. | Extend the validation logic to recognize and verify R1C1‑style references in named ranges in addition to A1 notation.
// Common Searches: how to programmatically validate named range addresses in Excel with Aspose.Cells C# | C# Aspose.Cells verify that RefersTo strings follow A1 reference format | detect invalid named ranges when loading a workbook using Aspose.Cells for .NET | sample code to check A1 notation of defined names in an Excel file with Aspose.Cells
// Tags: named range A1 validation Aspose.Cells | CellArea.CreateCellArea address check | RefersTo property format verification .NET | iterate defined names workbook Aspose.Cells | Excel range string normalization C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates over all defined names, extracts each name’s RefersTo address, normalizes it, and uses CellArea.CreateCellArea to confirm the address follows A1 notation, reporting valid or invalid results before saving the workbook.
class NamedRangeValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Iterate through all defined names (named ranges) in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // The RefersTo property contains the address, e.g. "=Sheet1!$A$1:$B$5"
            string refersTo = namedRange.RefersTo;

            // Remove the leading '=' if present
            if (!string.IsNullOrEmpty(refersTo) && refersTo.StartsWith("="))
                refersTo = refersTo.Substring(1);

            // Extract the range part (ignore sheet name if present)
            string rangePart = refersTo;
            int exclIndex = refersTo.IndexOf('!');
            if (exclIndex >= 0 && exclIndex < refersTo.Length - 1)
                rangePart = refersTo.Substring(exclIndex + 1);

            // Remove absolute reference symbols ('$') to simplify validation
            rangePart = rangePart.Replace("$", "");

            // Validate the range using CellArea.CreateCellArea (throws if invalid)
            bool isValid = true;
            try
            {
                // Split start and end cells; if only one cell is present, use it for both
                string startCell, endCell;
                int colonIdx = rangePart.IndexOf(':');
                if (colonIdx >= 0)
                {
                    startCell = rangePart.Substring(0, colonIdx);
                    endCell = rangePart.Substring(colonIdx + 1);
                }
                else
                {
                    startCell = endCell = rangePart;
                }

                // This will throw if the address is not a valid A1 style reference
                CellArea.CreateCellArea(startCell, endCell);
            }
            catch
            {
                isValid = false;
            }

            // Output validation result
            Console.WriteLine($"{namedRange.Text}: {(isValid ? "Valid" : "Invalid")} A1 address");
        }

        // Save the workbook after processing
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
