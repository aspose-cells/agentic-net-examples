// Title: Copy only changed cell values from a source worksheet to a target worksheet using Aspose.Cells for .NET
// Common Searches: how to copy only modified cells from one Excel sheet to another using Aspose.Cells in C# | sync two worksheets by comparing cell values with Aspose.Cells .NET | iterate over cells and update target sheet when values differ Aspose.Cells C# example | determine max used range of two worksheets and synchronize data Aspose.Cells | C# code to copy changed cells between worksheets in an existing workbook
// Tags: copy changed cells Aspose.Cells C# | worksheet cell comparison Aspose.Cells | synchronize worksheets Excel .NET | iterate cells max data range Aspose.Cells | update target worksheet based on source values

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, obtains the 'Source' and 'Target' worksheets, calculates the larger used range, iterates each cell, and copies the source value to the target only when the values differ, then saves the synchronized workbook.
class WorksheetSynchronizer
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputWorkbook.xlsx";
            const string outputFile = "SynchronizedWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook containing both source and target worksheets
            Workbook workbook = new Workbook(inputFile);

            // Get the source and target worksheets by name (adjust names as needed)
            Worksheet sourceSheet = workbook.Worksheets["Source"];
            Worksheet targetSheet = workbook.Worksheets["Target"];

            if (sourceSheet == null || targetSheet == null)
            {
                Console.WriteLine("Source or Target worksheet not found in the workbook.");
                return;
            }

            // Determine the maximum rows and columns to iterate (using the larger of the two used ranges)
            int maxRows = Math.Max(sourceSheet.Cells.MaxDataRow, targetSheet.Cells.MaxDataRow);
            int maxCols = Math.Max(sourceSheet.Cells.MaxDataColumn, targetSheet.Cells.MaxDataColumn);

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRows; row++)
            {
                for (int col = 0; col <= maxCols; col++)
                {
                    // Get the source cell
                    Cell sourceCell = sourceSheet.Cells[row, col];

                    // Skip if the source cell has no value
                    if (sourceCell.Value == null)
                        continue;

                    // Get the corresponding target cell
                    Cell targetCell = targetSheet.Cells[row, col];

                    // If values differ, copy the source value to the target
                    if (!object.Equals(sourceCell.Value, targetCell.Value))
                    {
                        targetCell.PutValue(sourceCell.Value);
                    }
                }
            }

            // Save the workbook with the synchronized target worksheet
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook synchronized and saved as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
