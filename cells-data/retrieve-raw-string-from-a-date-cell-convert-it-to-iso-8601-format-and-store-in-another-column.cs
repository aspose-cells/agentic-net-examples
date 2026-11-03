// Title: Extract dates from an Excel column, convert to ISO 8601 strings, and write them to the adjacent column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read date cell and output ISO 8601 string in another column | convert Excel DateTime cell to ISO 8601 format with Aspose.Cells .NET | iterate rows and write formatted date strings to adjacent column using Aspose.Cells | how to check CellValueType.IsDateTime and format as ISO 8601 in C# | save workbook after converting dates to ISO strings with Aspose.Cells
// Tags: read Excel date cells Aspose.Cells | format DateTime as ISO 8601 C# | write ISO string to neighboring column Aspose.Cells | CellValueType.IsDateTime usage | save modified workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads (or creates) an Excel workbook, scans each row in column A, detects cells that contain DateTime values, converts those dates to ISO 8601 strings, writes the strings into column B of the same rows, and finally saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a simple template.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                var ws = workbook.Worksheets[0];
                ws.Cells[0, 0].PutValue(DateTime.Now);
                ws.Cells[1, 0].PutValue(new DateTime(2023, 1, 15));
                workbook.Save(inputPath);
            }

            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            int sourceColumn = 0; // Column A
            int targetColumn = 1; // Column B

            int maxRow = cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                var sourceCell = cells[row, sourceColumn];

                // Process only cells recognized as DateTime
                if (sourceCell.Type == CellValueType.IsDateTime)
                {
                    DateTime date = sourceCell.DateTimeValue;
                    string isoString = date.ToString("o"); // ISO 8601 format
                    cells[row, targetColumn].PutValue(isoString);
                }
            }

            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
