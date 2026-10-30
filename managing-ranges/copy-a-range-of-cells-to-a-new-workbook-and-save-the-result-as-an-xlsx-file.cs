// Title: Copy a selected cell range from an existing workbook to a new workbook and save it as XLSX using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy A1:C10 to a new workbook while preserving formatting | How to duplicate a range of cells into a separate XLSX file with Aspose.Cells | C# example for copying a block of cells from one Excel file to another using Aspose.Cells | Preserve cell styles when copying a range to a new workbook in Aspose.Cells .NET | Create a new Excel file from a specific range of an existing file with Aspose.Cells
// Tags: copy cell range Aspose.Cells | preserve cell styles Aspose.Cells | create new workbook from range C# | save workbook as XLSX Aspose.Cells | transfer Excel range between workbooks .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source XLSX file, defines a cell block (e.g., A1:C10), creates a new workbook, copies each cell's value and style from the source to the destination sheet, and saves the result as a new XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string resultPath = "result.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (e.g., A1:C10)
            int startRow = 0;      // Row 1 (zero‑based)
            int startColumn = 0;   // Column A (zero‑based)
            int endRow = 9;        // Row 10
            int endColumn = 2;     // Column C

            int rowCount = endRow - startRow + 1;
            int columnCount = endColumn - startColumn + 1;

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Copy the defined range from the source sheet to the destination sheet
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    // Copy individual cell value and style
                    Cell sourceCell = sourceSheet.Cells[startRow + i, startColumn + j];
                    Cell destCell = destinationSheet.Cells[i, j];
                    destCell.PutValue(sourceCell.Value);
                    destCell.SetStyle(sourceCell.GetStyle());
                }
            }

            // Save the new workbook as an XLSX file
            destinationWorkbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Result workbook saved to: {resultPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
