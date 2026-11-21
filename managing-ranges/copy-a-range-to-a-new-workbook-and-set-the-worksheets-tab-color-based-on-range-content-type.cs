// Title: Copy a specific cell range to a new workbook and apply worksheet tab color based on numeric or text content using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy a specific cell range into a new workbook | How to change Excel worksheet tab color based on cell content type with Aspose.Cells | Determine if a range contains only numbers or only text using Aspose.Cells .NET | Copy cells with styles manually using Aspose.Cells C# example | Set worksheet tab color conditionally after copying data in Aspose.Cells
// Tags: range copy to separate workbook Aspose.Cells | conditional worksheet tab coloring Aspose.Cells | numeric versus string cell detection Aspose.Cells | preserve cell formatting during manual copy Aspose.Cells | C# Aspose.Cells content‑type based tab color

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads Source.xlsx, manually copies the A1:C10 range cell‑by‑cell (including styles) into a new workbook, checks whether the range contains only numeric values, only text, or a mix, and sets the destination worksheet tab color to green, blue, or yellow accordingly before saving as Destination.xlsx.
class Program
{
    static void Main()
    {
        const string sourcePath = "Source.xlsx";
        const string destinationPath = "Destination.xlsx";

        try
        {
            // Verify source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (example: A1:C10)
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 9,        // Row 10
                EndColumn = 2      // Column C
            };

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationWorksheet = destinationWorkbook.Worksheets[0];

            // Calculate rows and columns count for the range
            int totalRows = sourceArea.EndRow - sourceArea.StartRow + 1;
            int totalColumns = sourceArea.EndColumn - sourceArea.StartColumn + 1;

            // Manually copy the defined range from source to destination
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalColumns; c++)
                {
                    int srcRow = sourceArea.StartRow + r;
                    int srcCol = sourceArea.StartColumn + c;
                    int destRow = sourceArea.StartRow + r;
                    int destCol = sourceArea.StartColumn + c;

                    Cell srcCell = sourceWorksheet.Cells[srcRow, srcCol];
                    Cell destCell = destinationWorksheet.Cells[destRow, destCol];

                    // Copy value
                    destCell.Value = srcCell.Value;

                    // Copy style (if needed)
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Analyze the content type of the copied range
            bool containsNumber = false;
            bool containsString = false;

            for (int row = sourceArea.StartRow; row <= sourceArea.EndRow; row++)
            {
                for (int col = sourceArea.StartColumn; col <= sourceArea.EndColumn; col++)
                {
                    Cell cell = sourceWorksheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                        containsNumber = true;
                    else if (cell.Type == CellValueType.IsString)
                        containsString = true;
                }
            }

            // Set the worksheet tab color based on the detected content type
            if (containsNumber && !containsString)
                destinationWorksheet.TabColor = Color.Green;   // Only numeric values
            else if (containsString && !containsNumber)
                destinationWorksheet.TabColor = Color.Blue;    // Only text values
            else
                destinationWorksheet.TabColor = Color.Yellow;  // Mixed or other types

            // Ensure the directory for the destination file exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Save the new workbook
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Workbook saved successfully to '{destinationPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
