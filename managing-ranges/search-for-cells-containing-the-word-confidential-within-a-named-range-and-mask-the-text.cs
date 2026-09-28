// Title: Mask the word “confidential” in cells of a named range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook, retrieves a specific named range, and replaces every occurrence of the word “confidential” with asterisks using Aspose.Cells. | Generate a case‑insensitive regex routine that scans string cells inside a defined name range and masks the matched word with asterisks in C# with Aspose.Cells. | Create robust error‑handling for missing input files and undefined named ranges while performing text masking in an Excel worksheet using Aspose.Cells.
// Common Searches: aspocells c# mask confidential word in named range | replace specific text in defined name range using Aspose.Cells .NET | case insensitive search and replace in Excel cells with Aspose.Cells C#
// Tags: mask confidential text in named range Aspose.Cells | regex replace in Excel cells C# | defined name range iteration Aspose.Cells | case‑insensitive text masking workbook .NET | load and save workbook after text replacement Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file, obtains the named range "MyNamedRange", iterates through each string cell in that range, uses a case‑insensitive regular expression to find the word "confidential", replaces each occurrence with asterisks, and saves the modified workbook. It includes checks for missing files, undefined named ranges, and logs cell‑level errors without stopping the process.
class MaskConfidentialInNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string namedRangeName = "MyNamedRange";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the defined name (named range) from the workbook
            Name namedRange = workbook.Worksheets.Names[namedRangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{namedRangeName}\" not found.");
                return;
            }

            // Obtain the actual cell range that the name refers to
            AsposeRange range = namedRange.GetRange();

            // Case‑insensitive regex to locate the word "confidential"
            Regex confidentialRegex = new Regex("confidential", RegexOptions.IgnoreCase);

            // Iterate through each cell in the range
            Worksheet ws = range.Worksheet;
            for (int i = 0; i < range.RowCount; i++)
            {
                for (int j = 0; j < range.ColumnCount; j++)
                {
                    try
                    {
                        Cell cell = ws.Cells[range.FirstRow + i, range.FirstColumn + j];

                        // Process only string cells
                        if (cell.Value != null && cell.Type == CellValueType.IsString)
                        {
                            string originalText = cell.StringValue;

                            // Mask occurrences of the word "confidential"
                            if (confidentialRegex.IsMatch(originalText))
                            {
                                string maskedText = confidentialRegex.Replace(originalText,
                                    m => new string('*', m.Length));

                                cell.PutValue(maskedText);
                            }
                        }
                    }
                    catch (Exception cellEx)
                    {
                        // Log cell‑level errors but continue processing other cells
                        string colName = CellsHelper.ColumnIndexToName(range.FirstColumn + j);
                        Console.WriteLine($"Error processing cell {ws.Name}!{colName}{range.FirstRow + i + 1}: {cellEx.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
