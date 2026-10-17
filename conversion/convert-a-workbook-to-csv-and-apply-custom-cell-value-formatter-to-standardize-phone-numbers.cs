// Title: Convert an Excel workbook to CSV while normalizing phone numbers using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to CSV with custom cell formatting | How to reformat phone numbers in Excel before saving as CSV using Aspose.Cells | C# regex replace phone number patterns in workbook cells with Aspose.Cells | Save Excel file as CSV after standardizing US phone numbers in .NET | Iterate over worksheet cells to apply phone number format then convert to CSV with Aspose.Cells
// Tags: Aspose.Cells CSV export with custom cell formatter | C# regex phone number formatting in Excel worksheets | Iterate over worksheet cells to modify values Aspose.Cells | Excel to CSV conversion with data cleanup using .NET | Standardize US phone numbers during workbook export

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an .xlsx workbook with Aspose.Cells, scans all used cells, uses a regular expression to detect US‑style phone numbers, rewrites them to the (123) 456‑7890 format, and then saves the modified workbook as a CSV file.
class WorkbookToCsvWithPhoneFormatter
{
    static void Main()
    {
        // Define input and output file paths
        string inputFile = @"C:\Input\Workbook.xlsx";
        string outputFile = @"C:\Output\Workbook.csv";

        // Verify that the input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        // Ensure the output directory exists
        string outputDir = Path.GetDirectoryName(outputFile);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Regex to detect various phone number formats
            Regex phoneRegex = new Regex(@"\b(?:\+?1[-.\s]?)?(\(?\d{3}\)?)[-.\s]?(\d{3})[-.\s]?(\d{4})\b");

            // Determine the used range of cells
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Process only non‑empty string cells that look like phone numbers
                    if (cell.Type == CellValueType.IsString && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        string original = cell.StringValue;
                        Match match = phoneRegex.Match(original);
                        if (match.Success)
                        {
                            // Standardize to format: (123) 456-7890
                            string standardized = $"({match.Groups[1].Value.Trim('(', ')')}) {match.Groups[2].Value}-{match.Groups[3].Value}";
                            cell.PutValue(standardized);
                        }
                    }
                }
            }

            // Save the workbook as CSV (lifecycle rule: save)
            // Using SaveFormat.Csv; default options are sufficient for most scenarios.
            workbook.Save(outputFile, SaveFormat.Csv);

            Console.WriteLine($"Workbook successfully saved to CSV: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
