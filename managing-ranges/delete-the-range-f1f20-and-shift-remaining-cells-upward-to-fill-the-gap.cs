// Title: Delete rows F1:F20 and shift remaining cells upward using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to delete the rows covering the range F1:F20 and automatically shift the subsequent rows up. | Show how to create a range object for F1:F20 in Aspose.Cells, remove those rows, and save the workbook. | Provide a complete example that loads an Excel file, deletes the specified range, and writes the updated file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# delete rows F1 to F20 and shift up | How to remove a specific cell range and collapse rows in an Excel workbook with Aspose.Cells | C# Aspose.Cells DeleteRows method for range F1:F20 example | Delete range F1:F20 in Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells DeleteRows method C# | range deletion Aspose.Cells Excel | shift cells upward after row removal .NET | remove rows F1-F20 Aspose.Cells | Excel workbook manipulation Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a range for cells F1:F20 on the first worksheet, deletes the rows covered by that range (which automatically shifts the remaining rows upward), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Load existing workbook if the file exists; otherwise create a new workbook
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create the range F1:F20 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("F1", "F20");

            // Delete the rows covered by the range (shifts cells up automatically)
            int startRow = range.FirstRow;
            int rowsToDelete = range.RowCount;
            sheet.Cells.DeleteRows(startRow, rowsToDelete);

            // Path to the output workbook
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
