// Title: Copy a cell range to a new workbook and freeze the first row with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy range A1:D10 to a new workbook | How to freeze the top row after copying data with Aspose.Cells | Create a new workbook from a selected range using Aspose.Cells for .NET | Aspose.Cells FreezePanes example for first row in C# | Copy values and formatting between worksheets with Aspose.Cells | Save a copied cell range to a different file using Aspose.Cells C#
// Tags: range copy to separate workbook Aspose.Cells C# | freeze first row worksheet Aspose.Cells | manual cell style transfer Aspose.Cells | FreezePanes API usage Aspose.Cells .NET | export selected range to new file Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads a source XLSX file, manually copies the values and formatting of cells A1:D10 from the first worksheet into a newly created workbook, names the destination sheet "CopiedData", applies FreezePanes to lock the top row, ensures the output directory exists, and saves the result as a new XLSX file while handling missing files and runtime exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string sourcePath = "source.xlsx";
                string destinationPath = "dest.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);
                Worksheet srcWorksheet = srcWorkbook.Worksheets[0];

                // Create a new (empty) workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destWorksheet = destWorkbook.Worksheets[0];
                destWorksheet.Name = "CopiedData";

                // Define the range to copy (A1:D10)
                int startRow = 0;   // A1 -> row 0
                int startCol = 0;   // A1 -> column 0
                int endRow = 9;     // D10 -> row 9
                int endCol = 3;     // D10 -> column 3

                // Manually copy values and styles from source to destination
                for (int r = startRow; r <= endRow; r++)
                {
                    for (int c = startCol; c <= endCol; c++)
                    {
                        Cell srcCell = srcWorksheet.Cells[r, c];
                        Cell destCell = destWorksheet.Cells[r, c];

                        // Copy value
                        destCell.PutValue(srcCell.Value);

                        // Copy style
                        Style srcStyle = srcCell.GetStyle();
                        destCell.SetStyle(srcStyle);
                    }
                }

                // Freeze the top row in the destination worksheet
                // FreezePanes(row, column, totalRows, totalColumns)
                destWorksheet.FreezePanes(1, 0, 1, 0); // freezes the first row

                // Ensure the directory for the destination file exists
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the new workbook to a file
                destWorkbook.Save(destinationPath);
                Console.WriteLine($"Destination workbook saved to {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
