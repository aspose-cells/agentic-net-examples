// Title: How to import data from one Excel workbook into a named range of another workbook and overwrite existing cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy used range from source.xlsx into named range MyRange of target.xlsx | C# clear contents of a named range before writing external workbook data with Aspose.Cells | Overwrite existing cells in a named range with data from another workbook using Aspose.Cells .NET | Load source workbook and populate target workbook named range programmatically in C# | Aspose.Cells example for importing external workbook values into a specific named range
// Tags: Aspose.Cells import range from external workbook | overwrite named range cells Aspose.Cells | clear named range contents C# | populate Excel named range programmatically | copy used range to named range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads a source Excel file, extracts its used range, opens a target workbook, resolves the named range "MyRange", clears any existing data in that range, writes the source values cell‑by‑cell into the target range, and saves the updated workbook, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the source and target workbooks
            string sourcePath = "Source.xlsx";
            string targetPath = "Target.xlsx";
            string outputPath = "Target_Updated.xlsx";

            // Verify that input files exist
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target file not found: {targetPath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWb = new Workbook(sourcePath);
            Worksheet sourceWs = sourceWb.Worksheets[0];
            AsposeRange sourceRange = sourceWs.Cells.MaxDisplayRange;
            if (sourceRange == null)
            {
                Console.WriteLine("No data range found in the source workbook.");
                return;
            }

            object[,] sourceData = sourceRange.Value as object[,];
            if (sourceData == null)
            {
                Console.WriteLine("No data found in the source workbook.");
                return;
            }

            // Load the target workbook
            Workbook targetWb = new Workbook(targetPath);

            // Retrieve the named range from the target workbook
            Name namedRange = targetWb.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found in the target workbook.");
                return;
            }

            // Get the range object that represents the named range
            AsposeRange targetRange = namedRange.GetRange();
            if (targetRange == null)
            {
                Console.WriteLine("Unable to resolve the named range to a range object.");
                return;
            }

            // Determine the start cell of the target range
            int startRow = targetRange.FirstRow;
            int startColumn = targetRange.FirstColumn;
            Worksheet targetWs = targetRange.Worksheet; // Worksheet that contains the named range

            // Clear existing data in the target range
            targetWs.Cells.ClearContents(startRow, startColumn, targetRange.RowCount, targetRange.ColumnCount);

            // Write source data into the target range cell by cell
            int rows = sourceData.GetLength(0);
            int cols = sourceData.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    targetWs.Cells[startRow + i, startColumn + j].PutValue(sourceData[i, j]);
                }
            }

            // Save the updated target workbook
            targetWb.Save(outputPath);
            Console.WriteLine($"Updated workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
