// Title: Unmerge the D4:G4 range and apply left horizontal alignment to each cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unmerge a specific range D4 to G4 | Set left alignment for cells after unmerging in Aspose.Cells .NET | C# code to split merged cells and align text left using Aspose.Cells | How to restore individual cell formatting after unmerge with Aspose.Cells | Unmerge merged cells and apply style to each cell in Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells unmerge range C# | Aspose.Cells set cell horizontal alignment | C# unmerge merged cells Excel Aspose | Aspose.Cells apply style after unmerge | Excel workbook modify merged cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing Excel workbook, creates a Range object for cells D4:G4, calls UnMerge to split the merged block, then iterates over each cell in that range to set its HorizontalAlignment to Left, and finally saves the updated workbook. It also includes checks for the input file and creates the output directory if needed.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged range D4:G4 (zero‑based indexes)
            int startRow = 3;      // Row 4 (0‑based)
            int startColumn = 3;   // Column D (0‑based)
            int totalRows = 1;
            int totalColumns = 4;  // D to G inclusive

            // Unmerge the specified block using a Range object
            AsposeRange range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
            range.UnMerge();

            // Restore left alignment for each cell in the former merged block
            for (int col = startColumn; col < startColumn + totalColumns; col++)
            {
                Cell cell = sheet.Cells[startRow, col];
                Style style = cell.GetStyle();
                style.HorizontalAlignment = TextAlignmentType.Left;
                cell.SetStyle(style);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
