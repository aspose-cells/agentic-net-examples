// Title: Import values from a source Excel workbook into a named range of another workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy values from one workbook to a named range in another workbook | clear existing data and overwrite a named range with data from another Excel file using Aspose.Cells | create a matching sized range and transfer cell values between workbooks in .NET | load source and destination workbooks and copy named range values with Aspose.Cells | how to safely replace data in a named range from another workbook using Aspose.Cells for .NET
// Tags: copy values between named ranges Aspose.Cells | clear and overwrite named range .NET | create range with matching dimensions Aspose.Cells | import workbook data into specific range C# | transfer cells across workbooks Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source and a destination workbook, retrieves the named ranges "SourceRange" and "TargetRange", clears the destination range, creates a new range with the same dimensions as the source, copies each cell's value from the source range to the destination range, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string sourcePath = "Source.xlsx";
            string destinationPath = "Destination.xlsx";

            // Ensure input files exist
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");
            if (!File.Exists(destinationPath))
                throw new FileNotFoundException($"Destination file not found: {destinationPath}");

            // Load workbooks
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Workbook destinationWorkbook = new Workbook(destinationPath);

            // Retrieve the named range "TargetRange" from the destination workbook
            Aspose.Cells.Range destinationRange = destinationWorkbook.Worksheets.GetRangeByName("TargetRange");
            if (destinationRange == null)
                throw new InvalidOperationException("Named range 'TargetRange' not found in destination workbook.");

            // Clear existing data in the destination range
            destinationRange.ClearContents();

            // Retrieve the named range "SourceRange" from the source workbook
            Aspose.Cells.Range sourceRange = sourceWorkbook.Worksheets.GetRangeByName("SourceRange");
            if (sourceRange == null)
                throw new InvalidOperationException("Named range 'SourceRange' not found in source workbook.");

            // Determine size of the source data
            int rowCount = sourceRange.RowCount;
            int columnCount = sourceRange.ColumnCount;

            // Top‑left cell of the destination range
            int startRow = destinationRange.FirstRow;
            int startColumn = destinationRange.FirstColumn;

            // Create a target range in the destination worksheet with matching dimensions
            Worksheet destSheet = destinationRange.Worksheet;
            Aspose.Cells.Range targetRange = destSheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);

            // Copy values cell by cell
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    targetRange[i, j].Value = sourceRange[i, j].Value;
                }
            }

            // Save the updated destination workbook
            string outputPath = "Destination_Updated.xlsx";
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Updated workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
