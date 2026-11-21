// Title: Export a named range from an Excel workbook to a new .xlsx file while preserving cell formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export specific named range to a new workbook preserving formatting | How to copy a named range with styles to another Excel file using Aspose.Cells .NET | GetRangeByName example saving range as a separate .xlsx in C# | C# code to extract named range MyRange to a new workbook with Aspose.Cells | Preserve cell formatting when exporting Excel named range with Aspose.Cells
// Tags: export named range Aspose.Cells .NET | copy range preserving styles C# | GetRangeByName to new workbook | save range as separate .xlsx file | Aspose.Cells range extraction example

using System;
using System.IO;
using Aspose.Cells;

namespace ExportNamedRange
{
    // The sample loads 'source.xlsx', retrieves the named range 'MyRange' using GetRangeByName, copies the range—including all formatting—to a new workbook starting at cell A1, creates the destination folder if needed, and saves the result as 'exported.xlsx', with error handling for missing files and other exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string destinationPath = "exported.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);

                // Retrieve the named range "MyRange"
                // Use GetRangeByName which returns Aspose.Cells.Range
                Aspose.Cells.Range namedRange = srcWorkbook.Worksheets.GetRangeByName("MyRange");
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found in the source workbook.");
                    return;
                }

                // Determine dimensions of the named range
                int rowCount = namedRange.RowCount;
                int columnCount = namedRange.ColumnCount;

                // Create a new workbook for the exported range
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Create destination range starting at A1 with same size
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange(0, 0, rowCount, columnCount);

                // Copy the named range to the destination range (default copies everything)
                namedRange.Copy(destRange);

                // Ensure the directory for the destination file exists
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the exported workbook
                destWorkbook.Save(destinationPath);
                Console.WriteLine($"Exported range saved to '{destinationPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
