// Title: Replace a placeholder string in every cell of an Excel workbook using Aspose.Cells for .NET and regular expressions
// Common Searches: C# Aspose.Cells replace {PLACEHOLDER} in all cells of an Excel file using regex | How to perform bulk find and replace with a regular expression in an Aspose.Cells workbook | Iterate through every worksheet and substitute placeholder text with Aspose.Cells .NET | Using Aspose.Cells to replace placeholder strings across an entire workbook in C# | Regex based find and replace for cell values in Excel with Aspose.Cells library
// Tags: regex replace cell values Aspose.Cells | bulk placeholder substitution Excel C# | iterate worksheets replace text Aspose.Cells | find and replace across workbook Aspose.Cells | Aspose.Cells replace placeholder using regular expression

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, compiles a regular expression for the placeholder "{PLACEHOLDER}", iterates through each worksheet and every used cell, replaces matching strings with "ActualValue", and saves the modified workbook, including basic file‑existence checks and error handling.
class WorkbookPlaceholderReplacer
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Define the regular expression pattern for the placeholder
        string placeholderPattern = @"\{PLACEHOLDER\}";
        Regex regex = new Regex(placeholderPattern, RegexOptions.Compiled);

        // Define the replacement text
        string replacementText = "ActualValue";

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range to limit the iteration
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Work only with cells that contain a string value
                    if (cell.Type == CellValueType.IsString)
                    {
                        string originalText = cell.StringValue;
                        if (originalText == null) continue;

                        // Replace placeholder if it matches the pattern
                        if (regex.IsMatch(originalText))
                        {
                            string newText = regex.Replace(originalText, replacementText);
                            cell.PutValue(newText);
                        }
                    }
                }
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
