// Title: Convert a specified cell range into a structured ListObject table in an existing Excel workbook with Aspose.Cells for .NET
// Common Searches: how to add a ListObject table to an existing Excel file using Aspose.Cells C# | Aspose.Cells create structured table from a range of cells programmatically | C# command line tool to convert a cell range into an Excel table with Aspose | adding a header‑row table to a worksheet via Aspose.Cells API | save workbook after inserting ListObject table using Aspose.Cells .NET
// Tags: Aspose.Cells ListObject table creation | convert cell range to structured table .NET | add table with header row Aspose.Cells | programmatic workbook modification Aspose.Cells | command‑line Excel table insertion C# | save workbook after ListObject addition

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing Excel workbook, validates command‑line arguments for input and output paths and a zero‑based cell range, adds a ListObject (structured table) over that range with a header row, ensures the output directory exists, and saves the modified workbook to the specified location.
class ConvertListToTable
{
    static void Main(string[] args)
    {
        // Expected arguments:
        // 0 - input workbook path
        // 1 - output workbook path
        // 2 - start row (zero‑based)
        // 3 - start column (zero‑based)
        // 4 - end row (zero‑based)
        // 5 - end column (zero‑based)
        if (args.Length < 6)
        {
            Console.WriteLine("Usage: ConvertListToTable <inputPath> <outputPath> <startRow> <startColumn> <endRow> <endColumn>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Validate input file existence
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        // Parse range indices
        if (!int.TryParse(args[2], out int startRow) ||
            !int.TryParse(args[3], out int startColumn) ||
            !int.TryParse(args[4], out int endRow) ||
            !int.TryParse(args[5], out int endColumn))
        {
            Console.WriteLine("Error: Start/End row and column values must be valid integers.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Add a ListObject (structured table) over the specified range.
            // The last parameter 'true' indicates that the range includes a header row.
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Optional: apply a built‑in table style if needed.
            // Aspose.Cells may not expose TableStyles in all versions; this block is omitted for compatibility.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
