// Title: Remove asterisk characters from all string cells in a specific named range using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans a named range and deletes every "*" character from string cells. | Adapt the sample to replace asterisks with a custom substring while iterating over a named range in an Excel workbook. | Add comprehensive error handling for missing files and undefined named ranges when cleaning characters in a range.
// Common Searches: aspocells c# remove asterisk from cells in a defined name range | how to clean special characters in a named range using Aspose.Cells .NET | replace * character in Excel named range programmatically with Aspose.Cells | iterate through cells of a named range and update string values in C#
// Tags: remove asterisk Aspose.Cells named range | string cleanup in Excel range C# | Aspose.Cells replace characters in defined name | iterate cells in named range Aspose.Cells | error handling missing workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, retrieves the named range "MyRange", removes all asterisk characters from string cells within that range, and saves the updated file.
class ReplaceAsteriskInNamedRange
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string rangeName = "MyRange";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range by its defined name
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(rangeName);
            if (range == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string original = cell.StringValue;
                    string replaced = original.Replace("*", string.Empty);
                    // Update cell only if a change occurred
                    if (!original.Equals(replaced))
                    {
                        cell.PutValue(replaced);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Asterisk characters removed from named range \"{rangeName}\" and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
