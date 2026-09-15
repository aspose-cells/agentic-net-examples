// Title: Replace a {{PLACEHOLDER}} token with a specific value in every cell of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, scans all worksheets' used ranges, and uses a case‑insensitive regular expression to replace every occurrence of {{PLACEHOLDER}} with a given string, then saves the workbook. | Create a reusable C# method that takes a workbook path, a regex pattern, and replacement text, iterates through all string cells across all sheets via Aspose.Cells, applies the regex replace, and writes the updated file.
// Common Searches: aspocells c# replace placeholder token in all worksheets | how to perform regex find and replace across an entire Excel workbook with Aspose.Cells | bulk replace {{PLACEHOLDER}} in Excel using Aspose.Cells .NET | iterate used range cells and replace text with regex Aspose.Cells C# | case‑insensitive placeholder substitution in Excel file via Aspose.Cells
// Tags: regex find replace Aspose.Cells workbook | bulk cell text substitution .NET | replace placeholder in Excel using Aspose.Cells | case-insensitive text replace Aspose.Cells C# | iterate worksheets used range Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an input.xlsx workbook with Aspose.Cells, defines a case‑insensitive regex for the {{PLACEHOLDER}} token, walks through each worksheet's used range, replaces matching string cells with the specified value, and saves the result as output.xlsx while handling missing files and runtime errors.
class ReplacePlaceholderInWorkbook
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Regular expression to match the placeholder {{PLACEHOLDER}} (case‑insensitive).
            string pattern = @"\{\{PLACEHOLDER\}\}";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

            // Replacement text.
            const string replacement = "ActualValue";

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet.
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Loop through each cell within the used range.
                for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string cells.
                        if (cell.Type == CellValueType.IsString)
                        {
                            string originalText = cell.StringValue;
                            string newText = regex.Replace(originalText, replacement);

                            // Update the cell only if the text has changed.
                            if (!originalText.Equals(newText))
                            {
                                cell.PutValue(newText);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
