// Title: Copy rows values only from one Excel workbook to another using Aspose.Cells for .NET (ignore formulas and formatting)
// Common Searches: asp.net copy rows from source workbook to destination workbook without formulas using Aspose.Cells | c# copy only cell values from Excel sheet with Aspose.Cells | how to ignore cell formatting when copying rows in Aspose.Cells for .NET | copy multiple rows values only Aspose.Cells C# example | transfer Excel data values only using Aspose.Cells PutValue method
// Tags: row value copy Aspose.Cells .NET | ignore formulas Aspose.Cells C# | skip formatting during Excel row transfer Aspose.Cells | PutValue method for copying cell values Aspose.Cells | bulk row copy values only Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source workbook, creates a new destination workbook, and iterates over the first ten rows. For each cell it uses PutValue to transfer only the raw value, deliberately skipping formulas and any formatting, then saves the destination file.
class Program
{
    static void Main()
    {
        try
        {
            string sourcePath = "source.xlsx";
            string destinationPath = "destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWb = new Workbook(sourcePath);
            Worksheet srcWs = srcWb.Worksheets[0];

            // Create a new destination workbook
            Workbook destWb = new Workbook();
            Worksheet destWs = destWb.Worksheets[0];

            // Define the rows to copy (example: rows 0‑9)
            int sourceStartRow = 0;
            int destinationStartRow = 0;
            int rowCount = 10;

            // Copy values only (ignore formulas and formatting)
            for (int i = 0; i < rowCount; i++)
            {
                int srcRow = sourceStartRow + i;
                int destRow = destinationStartRow + i;

                // Determine the number of columns with data in the source sheet
                int maxColumn = srcWs.Cells.MaxColumn + 1; // MaxColumn is zero‑based

                for (int col = 0; col < maxColumn; col++)
                {
                    Cell srcCell = srcWs.Cells[srcRow, col];
                    Cell destCell = destWs.Cells[destRow, col];

                    // Transfer only the cell value
                    destCell.PutValue(srcCell.Value);
                }
            }

            // Save the destination workbook
            destWb.Save(destinationPath);
            Console.WriteLine($"Rows copied successfully to '{destinationPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
