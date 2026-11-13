// Title: Copy only cell formatting between worksheets using PasteOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy cell formatting without values using PasteOptions | How to use PasteType.Formats to transfer only styles between worksheets in .NET | Copy row and column formatting but exclude formulas with Aspose.Cells | PasteOptions example for formatting-only copy in C# Aspose.Cells | Transfer Excel cell styles to another sheet while keeping data unchanged using Aspose.Cells
// Tags: PasteOptions style-only transfer Aspose.Cells | Transfer cell styles between worksheets C# | PasteType.Formats implementation Aspose.Cells | Excel range style copy .NET | Range.Copy without data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a source workbook, creates or retrieves a target worksheet, defines a source range (A1:D10) and a matching destination range, configures PasteOptions with PasteType.Formats to copy only formatting, copies the formatting from the source range to the destination range, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string sourcePath = "Source.xlsx";
            const string resultPath = "Result.xlsx";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"The source file '{sourcePath}' was not found.");

                // Load the existing workbook
                var workbook = new Workbook(sourcePath);

                // Get the first worksheet (source)
                var sourceSheet = workbook.Worksheets[0];

                // Get or create the target worksheet
                Worksheet targetSheet = workbook.Worksheets["FormattedCopy"];
                if (targetSheet == null)
                    targetSheet = workbook.Worksheets.Add("FormattedCopy");

                // Define the range of rows/columns to copy (e.g., A1:D10)
                var sourceRange = sourceSheet.Cells.CreateRange("A1:D10");

                // Create a destination range of the same size starting at A1 in the target sheet
                var destRange = targetSheet.Cells.CreateRange(
                    firstRow: 0,
                    firstColumn: 0,
                    totalRows: sourceRange.RowCount,
                    totalColumns: sourceRange.ColumnCount);

                // Configure PasteOptions to copy only formatting
                var pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.Formats
                };

                // Copy the formatting from sourceRange to the destination range
                sourceRange.Copy(destRange, pasteOptions);

                // Save the workbook with the copied formatting
                workbook.Save(resultPath);
                Console.WriteLine($"Formatting copied successfully. Result saved to '{resultPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
