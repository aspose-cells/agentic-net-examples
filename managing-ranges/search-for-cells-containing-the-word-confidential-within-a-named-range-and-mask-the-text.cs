// Title: How to mask the word “confidential” in cells of a named range using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells replace confidential text in a named range | search for confidential keyword and mask it in Excel using Aspose.Cells .NET | iterate over cells in a named range and hide sensitive words with Aspose.Cells | case‑insensitive string replacement in a named range with Aspose.Cells C# | mask sensitive words in an Excel workbook programmatically using Aspose.Cells
// Tags: named range cell iteration Aspose.Cells C# | regex based text masking Excel Aspose.Cells | case‑insensitive word replace Aspose.Cells | sensitive data redaction Excel C# | replace confidential keyword Aspose.Cells workbook

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, retrieves a named range, iterates through each cell in that range, detects the word “confidential” (case‑insensitive) in string cells, replaces it with asterisks using a regular expression, writes the masked text back, and saves the modified workbook.
class MaskConfidentialInNamedRange
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string namedRangeName = "MyNamedRange";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            AsposeRange range = workbook.Worksheets.GetRangeByName(namedRangeName);
            if (range == null)
            {
                Console.WriteLine($"Named range \"{namedRangeName}\" not found.");
                return;
            }

            // Get the worksheet that contains the range
            Worksheet sheet = range.Worksheet;

            // Iterate through all cells in the named range
            for (int r = 0; r < range.RowCount; r++)
            {
                for (int c = 0; c < range.ColumnCount; c++)
                {
                    try
                    {
                        Cell cell = sheet.Cells[range.FirstRow + r, range.FirstColumn + c];

                        // Process only string cells
                        if (cell.Value != null && cell.Type == CellValueType.IsString)
                        {
                            string cellText = cell.StringValue;
                            // Check for the word "confidential" (case‑insensitive)
                            if (cellText.IndexOf("confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Mask the word "confidential" with asterisks while preserving other text
                                string maskedText = Regex.Replace(
                                    cellText,
                                    "(confidential)",
                                    new string('*', "confidential".Length),
                                    RegexOptions.IgnoreCase);

                                // Write the masked text back to the cell
                                cell.PutValue(maskedText);
                            }
                        }
                    }
                    catch (Exception cellEx)
                    {
                        Console.WriteLine($"Error processing cell at row {range.FirstRow + r}, column {range.FirstColumn + c}: {cellEx.Message}");
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
