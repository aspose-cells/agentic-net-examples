// Title: Merge two external Excel files into a single workbook while preserving styles using Aspose.Cells for .NET (C#)
// Common Searches: asp.net merge two Excel workbooks into one with Aspose.Cells C# | copy data from multiple Excel files and keep formatting using Aspose.Cells | C# program to combine source Excel files into a merged workbook Aspose.Cells | how to concatenate Excel workbooks programmatically with Aspose.Cells .NET | load two workbooks and append their data to a new workbook using Aspose.Cells C#
// Tags: merge excel workbooks Aspose.Cells C# | copy worksheet data preserving styles Aspose.Cells | append external excel source to new workbook Aspose.Cells | load multiple workbooks combine cells Aspose.Cells | create merged workbook from source files C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for two source Excel files, loads each workbook with Aspose.Cells, copies the used range of the first worksheet from both sources into a new workbook while preserving cell values and styles, appends the second source data after the first, creates the output directory if needed, and saves the merged workbook as 'MergedDataWorkbook.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to external source files
            string source1Path = @"C:\Data\Source1.xlsx";
            string source2Path = @"C:\Data\Source2.xlsx";

            // Verify that source files exist
            if (!File.Exists(source1Path) || !File.Exists(source2Path))
            {
                Console.WriteLine("One or both source files are missing. Please verify the paths.");
                return;
            }

            // Load source workbooks
            Workbook sourceWb1;
            Workbook sourceWb2;
            try
            {
                sourceWb1 = new Workbook(source1Path);
                sourceWb2 = new Workbook(source2Path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load source workbooks: {ex.Message}");
                return;
            }

            // Create a new workbook for merged data
            Workbook mergedWb = new Workbook();
            Worksheet mergedSheet = mergedWb.Worksheets[0];

            // Helper to copy data from a source worksheet to the merged sheet
            void CopyData(Worksheet srcSheet, int startRow)
            {
                // Determine used range
                var srcRange = srcSheet.Cells.MaxDisplayRange;
                int rowCount = srcRange.RowCount;
                int colCount = srcRange.ColumnCount;

                // Copy cells
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        var srcCell = srcSheet.Cells[r, c];
                        var destCell = mergedSheet.Cells[startRow + r, c];
                        destCell.PutValue(srcCell.Value);
                        // Copy style if needed
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }
            }

            // Copy data from first source (starts at row 0)
            CopyData(sourceWb1.Worksheets[0], 0);

            // Determine next empty row after first source data
            int nextRow = mergedSheet.Cells.MaxDisplayRange.RowCount;

            // Copy data from second source (starts after first source)
            CopyData(sourceWb2.Worksheets[0], nextRow);

            // Save the merged workbook
            string outputPath = "MergedDataWorkbook.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            mergedWb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
