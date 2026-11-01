// Title: Replace dd/MM/yyyy and dd-MM-yyyy strings with ISO 8601 (yyyy‑MM‑dd) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to convert date strings dd/MM/yyyy to yyyy-MM-dd in an Excel file with Aspose.Cells C# | using regular expressions to reformat dates in Excel cells via Aspose.Cells .NET | iterate over used range and update string cells to ISO 8601 format with Aspose.Cells | replace date patterns in worksheet cells using Aspose.Cells and Regex in C# | convert mixed date formats to ISO 8601 in an .xlsx using Aspose.Cells library
// Tags: regex date reformatting Aspose.Cells | replace string cell values Excel .NET | maxdisplayrange iteration Aspose.Cells | ISO 8601 conversion Excel workbook C# | handle varied date patterns Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an input.xlsx workbook with Aspose.Cells, scans the used range of the first worksheet, uses a regular expression to find dates in dd/MM/yyyy or dd‑MM‑yyyy format, converts each match to ISO 8601 (yyyy‑MM‑dd) via DateTime, updates the cell only when a change occurs, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Regex pattern to match dates like dd/MM/yyyy or dd-MM-yyyy
            string pattern = @"\b(?<day>\d{2})[\/\-](?<month>\d{2})[\/\-](?<year>\d{4})\b";

            // Get the used range of the worksheet (fully qualified to avoid ambiguity)
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = startRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = startCol + usedRange.ColumnCount - 1;

            // Iterate through each cell in the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        string originalText = cell.StringValue;

                        // Replace matched dates with ISO 8601 format (yyyy-MM-dd)
                        string newText = Regex.Replace(originalText, pattern, match =>
                        {
                            int day = int.Parse(match.Groups["day"].Value);
                            int month = int.Parse(match.Groups["month"].Value);
                            int year = int.Parse(match.Groups["year"].Value);
                            DateTime dt = new DateTime(year, month, day);
                            return dt.ToString("yyyy-MM-dd");
                        });

                        // Update the cell only if a replacement occurred
                        if (!newText.Equals(originalText))
                        {
                            cell.PutValue(newText);
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
