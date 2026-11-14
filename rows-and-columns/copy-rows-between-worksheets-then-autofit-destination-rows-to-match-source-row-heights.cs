// Title: Copy a range of rows between two Excel worksheets and keep original row heights using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy rows from one worksheet to another preserving original row height | How to copy multiple rows with formatting and match row heights using Aspose.Cells .NET | Auto‑fit rows after copying rows between Excel workbooks with Aspose.Cells in C# | Copy rows 0 to 9 from source workbook to destination workbook Aspose.Cells example | Set destination row height equal to source row height Aspose.Cells C# code
// Tags: copy rows between worksheets Aspose.Cells | match source row height Aspose.Cells | auto‑fit destination rows Aspose.Cells | copy cell values and styles Aspose.Cells | load source and destination workbooks Aspose.Cells | set row height programmatically Aspose.Cells | copy rows with formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads a source and a destination workbook, copies rows 0‑9 from the first worksheet of the source to the first worksheet of the destination—including values, formulas, and styles—sets each destination row's Height to match the source row, auto‑fits any remaining rows, and saves the result as a new Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for source and destination workbooks
            string sourcePath = "Source.xlsx";
            string destinationPath = "Destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Load destination workbook if it exists; otherwise create a new one
            Workbook destinationWorkbook = File.Exists(destinationPath)
                ? new Workbook(destinationPath)
                : new Workbook();

            // Get the first worksheet from each workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the range of rows to copy (zero‑based indexes)
            int sourceStartRow = 0;      // first row to copy
            int sourceEndRow   = 9;      // last row to copy (inclusive)
            int destinationStartRow = 0; // where to start pasting in the destination sheet

            // Determine the number of columns to copy
            int maxColumn = Math.Max(sourceSheet.Cells.MaxColumn, destinationSheet.Cells.MaxColumn);

            // Copy rows one by one
            for (int srcRow = sourceStartRow, destRow = destinationStartRow;
                 srcRow <= sourceEndRow;
                 srcRow++, destRow++)
            {
                // Copy each cell in the row (value, formula, style, etc.)
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell srcCell = sourceSheet.Cells[srcRow, col];
                    Cell destCell = destinationSheet.Cells[destRow, col];

                    // Copy value (includes formulas)
                    destCell.PutValue(srcCell.Value);

                    // Copy style
                    destCell.SetStyle(srcCell.GetStyle());

                    // Additional properties (comments, hyperlinks, etc.) can be copied here if needed
                }

                // Match the row height from source to destination
                Row srcRowObj = sourceSheet.Cells.Rows[srcRow];
                Row destRowObj = destinationSheet.Cells.Rows[destRow];
                destRowObj.Height = srcRowObj.Height;
            }

            // Auto‑fit rows in the destination sheet for any rows not explicitly copied
            destinationSheet.AutoFitRows();

            // Save the destination workbook
            string outputPath = "Destination_Copied.xlsx";
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Rows copied successfully. Output saved to: {outputPath}");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
