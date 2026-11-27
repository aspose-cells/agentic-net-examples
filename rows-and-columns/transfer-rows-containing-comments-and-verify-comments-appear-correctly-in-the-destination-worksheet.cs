// Title: Transfer rows that contain Excel comments to a new workbook and verify them using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy only rows with comments to another workbook | how to preserve cell comments when copying rows with Aspose.Cells | verify transferred comments in destination worksheet using Aspose.Cells C# | extract rows that have Excel comments and save to new file with Aspose.Cells | Aspose.Cells retain comment author and note during row transfer
// Tags: duplicate rows containing comments via Aspose.Cells | retain comment author and text when moving Excel rows | select rows that have comments for .NET workbook export | check comment integrity after row migration in C# | Aspose.Cells CopyRows preserving comment metadata

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source XLSX file, scans each row for cell comments, copies rows that contain at least one comment (including values, formatting, and comment metadata) to a new worksheet, prints the transferred comments (author and note) for verification, and saves the result as a new workbook.
class TransferRowsWithComments
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string destinationPath = "Destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new destination workbook
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            destinationSheet.Name = "RowsWithComments";

            int destRowIndex = 0; // tracks the next free row in the destination sheet

            // Determine the used range in the source sheet
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = sourceSheet.Cells.MaxDataColumn;

            // Iterate through each row in the source sheet
            for (int srcRow = 0; srcRow <= maxRow; srcRow++)
            {
                bool rowContainsComment = false;

                // Check every cell in the current row for a comment
                for (int col = 0; col <= maxCol; col++)
                {
                    if (sourceSheet.Comments[srcRow, col] != null)
                    {
                        rowContainsComment = true;
                        break;
                    }
                }

                // If the row has at least one comment, copy it to the destination sheet
                if (rowContainsComment)
                {
                    // Copy the entire row (values, formats, etc.) from source to destination
                    destinationSheet.Cells.CopyRows(sourceSheet.Cells, srcRow, destRowIndex, 1);

                    // Copy comments cell‑by‑cell
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Comment srcComment = sourceSheet.Comments[srcRow, col];
                        if (srcComment != null)
                        {
                            // Create a new comment in the destination cell
                            Comment destComment = destinationSheet.Comments[destRowIndex, col];
                            if (destComment == null)
                            {
                                destComment = destinationSheet.Comments[destRowIndex, col];
                            }

                            // Copy comment properties
                            destComment.Note = srcComment.Note;
                            destComment.Author = srcComment.Author;
                        }
                    }

                    destRowIndex++; // move to the next free row in the destination sheet
                }
            }

            // Verification: output all comments found in the destination worksheet
            Console.WriteLine("Comments in the destination worksheet:");
            for (int row = 0; row < destRowIndex; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Comment comment = destinationSheet.Comments[row, col];
                    if (comment != null)
                    {
                        Console.WriteLine($"Row {row + 1}, Column {col + 1}: Author='{comment.Author}', Note='{comment.Note}'");
                    }
                }
            }

            // Ensure the directory for the destination file exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the destination workbook
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Destination workbook saved to: {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
