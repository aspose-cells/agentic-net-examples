// Title: Copy columns A through D from a source worksheet to a new workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# copy specific columns from one Excel file to another with Aspose.Cells | how to transfer columns A to D between workbooks using Aspose.Cells API | Aspose.Cells copy selected column range preserving styles in C# | efficiently copy first four columns to a new workbook without LightCells | Aspose.Cells example for copying column data and formatting in .NET
// Tags: Aspose.Cells copy column range C# | transfer selected columns between Excel workbooks | preserve cell styles during column copy Aspose.Cells | .NET Excel column extraction example | copy columns A-D using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a source workbook, creates a new workbook, iterates through each used row, and copies the values and styles of cells in columns A through D (indices 0‑3) from the source worksheet to the destination worksheet, then saves the new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source file \"{sourcePath}\" not found.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new workbook for the destination
                Workbook destinationWorkbook = new Workbook();

                // Get the first worksheet from the source
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Get the first worksheet of the destination workbook
                Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

                // Determine the number of rows to copy (including the last used row)
                int maxRow = sourceSheet.Cells.MaxDataRow;

                // Copy columns A (0) through D (3) from each row
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= 3; col++)
                    {
                        // Get the source cell
                        Cell sourceCell = sourceSheet.Cells[row, col];

                        // Get the corresponding destination cell
                        Cell destCell = destinationSheet.Cells[row, col];

                        // Copy value
                        destCell.PutValue(sourceCell.Value);

                        // Copy style (optional, can be omitted if only values are needed)
                        destCell.SetStyle(sourceCell.GetStyle());
                    }
                }

                // Save the new workbook
                destinationWorkbook.Save(destinationPath);
                Console.WriteLine($"Destination workbook saved to \"{destinationPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
