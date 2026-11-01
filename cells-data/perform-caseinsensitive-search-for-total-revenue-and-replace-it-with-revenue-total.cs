// Title: Perform case‑insensitive search and replace of 'total revenue' with 'Revenue Total' in all worksheets of an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace text case‑insensitively in entire workbook | How to find and replace a phrase in every sheet of an Excel file with Aspose.Cells .NET | Using Regex with Aspose.Cells to change 'total revenue' to 'Revenue Total' in Excel | Iterate through used range cells and update string values in Aspose.Cells C# | Case‑insensitive find and replace across multiple worksheets using Aspose.Cells API
// Tags: case-insensitive find replace Aspose.Cells | replace text in Excel workbook C# | iterate worksheets used range Aspose.Cells | regex text substitution spreadsheet cells | update string cell values Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an input.xlsx workbook with Aspose.Cells, walks through each worksheet and its used cell range, detects the phrase "total revenue" in string cells without regard to case, replaces every occurrence with "Revenue Total" using a case‑insensitive regular expression, writes the modified values back, and saves the result as output.xlsx while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the worksheet (Aspose.Cells returns a Range object)
                var usedRange = cells.MaxDisplayRange; // use var to avoid ambiguity with System.Range
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell within the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only string cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;

                            // Check for the target phrase case‑insensitively
                            if (!string.IsNullOrEmpty(text) &&
                                text.IndexOf("total revenue", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace all occurrences of 'total revenue' with 'Revenue Total' (case‑insensitive)
                                string replaced = Regex.Replace(
                                    text,
                                    "total revenue",
                                    "Revenue Total",
                                    RegexOptions.IgnoreCase);

                                // Write the new value back to the cell
                                cell.PutValue(replaced);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
