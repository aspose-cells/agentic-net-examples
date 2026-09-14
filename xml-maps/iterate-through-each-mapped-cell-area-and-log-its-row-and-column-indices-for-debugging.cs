// Title: Log row and column indices of each cell in a named XML-mapped range using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, retrieves a named range, and writes each cell's row and column numbers to the Debug output. | Show how to calculate the first and last row/column of a named XML-mapped range and iterate through the cells for logging in Aspose.Cells.
// Common Searches: how to debug cells in a named range with Aspose.Cells C# | Aspose.Cells get row and column of each cell in XML mapped range | enumerate cells of a named range and log coordinates using Aspose.Cells .NET | retrieve start and end indices of a named range in Excel with Aspose.Cells
// Tags: Aspose.Cells named range iteration | C# debug cell indices Aspose.Cells | XML map range logging Aspose.Cells | range start end calculation Aspose.Cells | cell coordinate extraction .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDebug
{
    // The example loads an Excel workbook, accesses the first worksheet, creates a range named "MyMappedRange", computes its start and end rows and columns, logs the overall range boundaries, iterates through each cell within the range writing its row and column indices to the Debug output, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load the workbook (implementation checks file existence).
                Workbook workbook = LoadWorkbook();

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the named range that represents the mapped cell areas.
                AsposeRange mappedRange = worksheet.Cells.CreateRange("MyMappedRange");

                // Determine the boundaries of the range.
                int startRow = mappedRange.FirstRow;
                int startColumn = mappedRange.FirstColumn;
                int endRow = startRow + mappedRange.RowCount - 1;
                int endColumn = startColumn + mappedRange.ColumnCount - 1;

                // Log the overall range boundaries.
                Debug.WriteLine(
                    $"Range StartRow={startRow}, StartColumn={startColumn}, " +
                    $"EndRow={endRow}, EndColumn={endColumn}");

                // Iterate through each cell within the range.
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Debug.WriteLine($"Cell at Row={row}, Column={col}");
                    }
                }

                // Save the workbook if needed.
                SaveWorkbook(workbook);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Debug.WriteLine($"Error: {ex.Message}");
            }
        }

        // Loads a workbook from a predefined file path, ensuring the file exists.
        static Workbook LoadWorkbook()
        {
            const string inputPath = "input.xlsx";

            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input workbook '{inputPath}' was not found.");

            // Load the workbook from the file.
            return new Workbook(inputPath);
        }

        // Saves the workbook to a predefined output file path.
        static void SaveWorkbook(Workbook workbook)
        {
            const string outputPath = "output.xlsx";

            try
            {
                workbook.Save(outputPath);
                Debug.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to save workbook: {ex.Message}");
                throw;
            }
        }
    }
}
