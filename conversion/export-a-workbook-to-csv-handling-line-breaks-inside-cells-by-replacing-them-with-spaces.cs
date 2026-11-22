// Title: Convert an Excel workbook to CSV and replace line breaks in cell text using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET replace newline characters in Excel cells before saving as CSV | How to remove CRLF from Excel cell values when exporting to CSV with C# | Save workbook as CSV with Aspose.Cells while cleaning multiline cell content | C# iterate through worksheets and replace line breaks in cells for CSV export | Aspose.Cells TxtSaveOptions CSV output without embedded line breaks
// Tags: replace newline characters in Excel cells Aspose.Cells | Aspose.Cells CSV export line break removal | TxtSaveOptions CSV UTF-8 encoding | iterate worksheets cells clean string values C# | export workbook to CSV without multiline cell content

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel file, walks through every worksheet and cell, replaces any carriage‑return or line‑feed characters in string cells with a space, and then saves the workbook as a UTF‑8 encoded CSV using Aspose.Cells TxtSaveOptions, while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and cells to replace line breaks with spaces
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        if (!string.IsNullOrEmpty(original) && (original.Contains("\n") || original.Contains("\r")))
                        {
                            // Replace CRLF, LF, and CR with a single space
                            string cleaned = original.Replace("\r\n", " ")
                                                    .Replace("\n", " ")
                                                    .Replace("\r", " ");
                            cell.PutValue(cleaned);
                        }
                    }
                }
            }

            // Save the workbook as CSV using TxtSaveOptions (compatible with all Aspose.Cells versions)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Encoding = System.Text.Encoding.UTF8
            };
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully saved as CSV to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
