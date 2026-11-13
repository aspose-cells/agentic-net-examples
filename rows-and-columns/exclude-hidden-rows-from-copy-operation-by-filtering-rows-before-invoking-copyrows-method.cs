// Title: Copy only visible rows from a source worksheet to a new sheet with Aspose.Cells for .NET (C#)
// Common Searches: how to copy only visible rows from one Excel sheet to another using Aspose.Cells C# | exclude hidden rows when duplicating worksheet data with Aspose.Cells for .NET | copy rows while preserving cell styles and skipping hidden rows in C# Aspose.Cells | filter rows by IsHidden before using CopyRows method in Aspose.Cells | duplicate Excel sheet without hidden rows using Aspose.Cells API
// Tags: copy visible rows Aspose.Cells C# | skip hidden rows worksheet copy .NET | preserve cell formatting during row copy Aspose.Cells | filter rows by IsHidden property C# | duplicate worksheet data without hidden rows Excel

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, creates a new worksheet, iterates through each row of the first sheet, checks the Row.IsHidden flag, and copies only the visible rows' values and styles to the new sheet before saving the result.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Source and destination worksheets
            Worksheet srcSheet = workbook.Worksheets[0];
            Worksheet destSheet = workbook.Worksheets.Add("Copy");

            // Determine the range of rows that contain data
            int startRow = 0;
            int endRow = srcSheet.Cells.MaxDataRow; // last row with data

            // Destination start row index
            int destRow = 0;

            // Maximum column index to copy (covers all used columns)
            int maxColumn = srcSheet.Cells.MaxDataColumn;

            // Iterate through each row in the source sheet
            for (int i = startRow; i <= endRow; i++)
            {
                // Check if the current row is hidden
                Row row = srcSheet.Cells.Rows[i];
                if (row != null && !row.IsHidden)
                {
                    // Copy each cell in the visible row to the destination sheet
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell srcCell = srcSheet.Cells[i, col];
                        if (srcCell != null)
                        {
                            Cell destCell = destSheet.Cells[destRow, col];
                            // Copy value
                            destCell.PutValue(srcCell.Value);
                            // Copy style
                            destCell.SetStyle(srcCell.GetStyle());
                        }
                    }
                    destRow++; // move to next destination row
                }
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
