// Title: C# code to mask the word “confidential” in all cells of a named range using Aspose.Cells
// Common Searches: how to replace a specific word in an Excel named range with Aspose.Cells C# | mask confidential text in cells of a named range using Aspose.Cells .NET | case‑insensitive string replacement in an Excel named range with Aspose.Cells | iterate through cells of a named range and modify their content in C# | Aspose.Cells find and mask keyword in a workbook named range
// Tags: mask keyword in named range Aspose.Cells | case‑insensitive string replacement Aspose.Cells C# | iterate over named range cells Aspose.Cells | regex based text masking Excel Aspose.Cells | search cells by keyword Aspose.Cells workbook

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves a named range called "MyNamedRange", iterates through each cell in that range, and for any string cell containing the word "confidential" (case‑insensitive) replaces it with ten asterisks using a regular expression, then saves the modified workbook.
class MaskConfidentialInNamedRange
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string namedRangeName = "MyNamedRange";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (use GetRangeByName)
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName(namedRangeName);
            if (namedRange == null)
            {
                Console.WriteLine($"Error: Named range \"{namedRangeName}\" not found.");
                return;
            }

            // Iterate through all cells in the named range
            for (int row = namedRange.FirstRow; row <= namedRange.FirstRow + namedRange.RowCount - 1; row++)
            {
                for (int col = namedRange.FirstColumn; col <= namedRange.FirstColumn + namedRange.ColumnCount - 1; col++)
                {
                    Cell cell = namedRange.Worksheet.Cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;

                        // Check for the word "confidential" (case‑insensitive)
                        if (text.IndexOf("confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Mask the word with asterisks
                            string masked = Regex.Replace(
                                text,
                                "confidential",
                                "**********",          // 10 asterisks as mask
                                RegexOptions.IgnoreCase);

                            // Write the masked text back to the cell
                            cell.PutValue(masked);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
