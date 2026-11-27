// Title: Replace MM/dd/yyyy strings with ISO 8601 (yyyy-MM-dd) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace date strings in Excel cells using regex | How to convert MM/dd/yyyy text to yyyy-MM-dd in an Excel file with Aspose.Cells | Iterate over used range and update string cells in an Aspose.Cells workbook | Regex pattern for reformatting dates in Aspose.Cells C# example | Save modified workbook after date format conversion with Aspose.Cells
// Tags: regex date reformat Aspose.Cells C# | replace string cell values Aspose.Cells | iterate used range worksheets Aspose.Cells | convert MMddyyyy to ISO8601 Excel Aspose.Cells | save workbook after cell updates Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, scans each worksheet's used range, and uses a regular expression to transform text dates from MM/dd/yyyy to ISO 8601 (yyyy-MM-dd) in string cells, then saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex pattern for dates in MM/dd/yyyy format and ISO 8601 replacement
            string pattern = @"\b(\d{2})/(\d{2})/(\d{4})\b";
            string replacement = "$3-$1-$2";

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Process only cells that contain string values
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        // Perform regex replacement
                        string updated = Regex.Replace(original, pattern, replacement);
                        // If a change occurred, write the new value back to the cell
                        if (!original.Equals(updated, StringComparison.Ordinal))
                        {
                            cell.PutValue(updated);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
