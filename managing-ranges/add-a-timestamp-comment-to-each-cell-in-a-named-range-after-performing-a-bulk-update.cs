// Title: Add timestamp comments to each cell of a named range after bulk updating values with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add comment with current timestamp to cells in a named range | How to insert comments into all cells of a bulk-updated range using Aspose.Cells | C# code to set values in a named range then add timestamp notes with Aspose.Cells | Iterate over Aspose.Cells named range to attach comment containing DateTime.Now | Saving workbook after adding per-cell comments to a named range in .NET
// Tags: bulk update named range Aspose.Cells | add timestamp comment Aspose.Cells | named range cell comments C# | Aspose.Cells GetRangeByName example | save workbook after modifying comments .NET

using System;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, retrieves a named range, fills the range with a constant value using a bulk 2‑dimensional array, then iterates over each cell in that range to add or update a comment containing the current timestamp, and finally saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string rangeName = "MyRange";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range from the workbook (not from the worksheet)
            AsposeRange range = workbook.Worksheets.GetRangeByName(rangeName);
            if (range == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Bulk update: set all cells in the range to a constant value
            object[,] bulkData = new object[range.RowCount, range.ColumnCount];
            for (int i = 0; i < range.RowCount; i++)
            {
                for (int j = 0; j < range.ColumnCount; j++)
                {
                    bulkData[i, j] = "Updated";
                }
            }
            range.Value = bulkData;

            // Add a timestamp comment to each cell in the named range
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            for (int r = range.FirstRow; r <= range.FirstRow + range.RowCount - 1; r++)
            {
                for (int c = range.FirstColumn; c <= range.FirstColumn + range.ColumnCount - 1; c++)
                {
                    try
                    {
                        Cell cell = sheet.Cells[r, c];
                        // Add or update comment with timestamp
                        int commentIndex = sheet.Comments.Add(cell.Name);
                        Comment comment = sheet.Comments[commentIndex];
                        comment.Author = "System";
                        comment.Note = $"Timestamp: {timestamp}";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to add comment to cell {r},{c}: {ex.Message}");
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
