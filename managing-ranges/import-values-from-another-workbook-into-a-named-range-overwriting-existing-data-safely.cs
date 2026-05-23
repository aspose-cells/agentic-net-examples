using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace ImportIntoNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceData.xlsx";
                const string destPath = "Destination.xlsx";
                const string outputPath = "Destination_Updated.xlsx";

                // Verify that source and destination files exist
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                if (!File.Exists(destPath))
                {
                    Console.WriteLine($"Destination file '{destPath}' not found.");
                    return;
                }

                // Load workbooks
                Workbook sourceWb = new Workbook(sourcePath);
                Workbook destWb = new Workbook(destPath);

                // Retrieve the named range from the destination workbook
                // Assume the named range is called "ImportRange"
                Name importName = destWb.Worksheets.Names["ImportRange"];
                if (importName == null)
                {
                    Console.WriteLine("Named range 'ImportRange' not found in destination workbook.");
                    return;
                }

                // Get the actual range object that the name refers to
                AsposeRange destRange = importName.GetRange();

                // Define the source range to copy from (first worksheet, cells A1:B10)
                Worksheet sourceSheet = sourceWb.Worksheets[0];
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:B10");

                // Ensure the source and destination ranges have the same dimensions
                if (sourceRange.RowCount != destRange.RowCount ||
                    sourceRange.ColumnCount != destRange.ColumnCount)
                {
                    Console.WriteLine("Source and destination ranges have different sizes. Adjusting destination range.");

                    // Clear the area that will be overwritten
                    destWb.Worksheets[destRange.Worksheet.Index].Cells.ClearRange(
                        destRange.FirstRow, destRange.FirstColumn,
                        sourceRange.RowCount, sourceRange.ColumnCount);

                    // Recreate destination range to match source size
                    destRange = destWb.Worksheets[destRange.Worksheet.Index].Cells.CreateRange(
                        destRange.FirstRow, destRange.FirstColumn,
                        sourceRange.RowCount, sourceRange.ColumnCount);
                }
                else
                {
                    // Clear existing contents in the destination range before copying
                    destWb.Worksheets[destRange.Worksheet.Index].Cells.ClearRange(
                        destRange.FirstRow, destRange.FirstColumn,
                        destRange.RowCount, destRange.ColumnCount);
                }

                // Copy data (including formulas, formatting, etc.) from source to destination
                destRange.CopyData(sourceRange);

                // Save the updated destination workbook
                destWb.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}