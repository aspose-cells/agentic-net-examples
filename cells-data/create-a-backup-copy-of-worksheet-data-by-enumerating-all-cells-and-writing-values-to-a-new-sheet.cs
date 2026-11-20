// Title: Backup an Excel worksheet by enumerating cells and copying values using Aspose.Cells for .NET (C#)
// Common Searches: C# code to copy all cell values from one worksheet to a new backup sheet with Aspose.Cells | How to enumerate used range cells and duplicate worksheet data in Aspose.Cells .NET | Aspose.Cells example for creating a backup worksheet without using the copy method | Iterate through cells and write values to another sheet in an Excel file using Aspose.Cells C# | Save a copy of worksheet data to a new workbook using Aspose.Cells for .NET
// Tags: enumerate cells with Aspose.Cells MaxDisplayRange | copy cell values to another worksheet Aspose.Cells | add backup worksheet programmatically Aspose.Cells | save workbook with new sheet Aspose.Cells .NET | backup Excel data using cell iteration Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing Excel file, adds a new worksheet named "Backup", determines the source sheet's used range via MaxDisplayRange, iterates through each cell in that range copying the cell's Value to the corresponding cell in the backup sheet, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook from a file
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet that contains the original data (e.g., the first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet that will hold the backup copy
            Worksheet backupSheet = workbook.Worksheets.Add("Backup");

            // Determine the used range of the source worksheet
            AsposeRange usedRange = sourceSheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startColumn = usedRange.FirstColumn;
            int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Enumerate each cell within the used range and copy its value to the backup sheet
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell sourceCell = sourceSheet.Cells[row, col];
                    Cell backupCell = backupSheet.Cells[row, col];

                    // Copy the cell's value (preserves numbers, strings, dates, etc.)
                    backupCell.Value = sourceCell.Value;
                }
            }

            // Save the workbook, now containing the backup sheet, to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
