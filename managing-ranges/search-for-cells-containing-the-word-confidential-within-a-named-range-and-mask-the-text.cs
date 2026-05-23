using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Parse the RefersTo string (e.g., "Sheet1!$A$1:$C$10")
            string refersTo = namedRange.RefersTo;
            int exclPos = refersTo.IndexOf('!');
            if (exclPos < 0)
            {
                Console.WriteLine("Invalid RefersTo format.");
                return;
            }

            string sheetName = refersTo.Substring(0, exclPos);
            string area = refersTo.Substring(exclPos + 1);

            // Get the worksheet that contains the range
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                Console.WriteLine($"Worksheet '{sheetName}' not found.");
                return;
            }

            // Create a Range object from the address
            AsposeRange range = sheet.Cells.CreateRange(area);

            // Set replace options: partial match, case‑insensitive
            ReplaceOptions replaceOptions = new ReplaceOptions
            {
                MatchEntireCellContents = false,
                CaseSensitive = false
            };

            // Iterate through all cells in the named range and mask the word "confidential"
            for (int row = range.FirstRow; row < range.FirstRow + range.RowCount; row++)
            {
                for (int col = range.FirstColumn; col < range.FirstColumn + range.ColumnCount; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsString && cell.Value != null)
                    {
                        cell.Replace("confidential", "***********", replaceOptions);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}