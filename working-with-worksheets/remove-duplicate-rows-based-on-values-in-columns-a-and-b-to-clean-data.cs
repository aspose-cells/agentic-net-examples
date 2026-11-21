// Title: Remove duplicate rows based on columns A and B in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove duplicate rows from an Excel file using Aspose.Cells | How to delete rows with duplicate values in columns A and B with Aspose.Cells .NET | Aspose.Cells RemoveDuplicates example for two columns in C# | Programmatically clean Excel data by eliminating duplicate rows in columns A and B | Using Aspose.Cells to deduplicate an Excel worksheet in a .NET application
// Tags: Aspose.Cells RemoveDuplicates C# | Excel duplicate row removal .NET | deduplicate worksheet columns A B Aspose.Cells | C# clean Excel data duplicates | Aspose.Cells worksheet deduplication

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, determines the used row range, and calls worksheet.Cells.RemoveDuplicates starting at row 0, column 0 for the first two columns (A and B) to eliminate duplicate rows. The cleaned workbook is then saved as a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the range that includes all used rows and columns A and B
            int firstRow = 0; // zero‑based index for the first row
            int firstColumn = 0; // column A (zero‑based)
            int totalRows = worksheet.Cells.MaxDataRow + 1; // total number of rows with data
            int totalColumns = 2; // columns A and B

            // Remove duplicate rows based on the values in columns A and B
            // The last parameter (columnOffsets) is required; null means no additional columns
            worksheet.Cells.RemoveDuplicates(firstRow, firstColumn, totalRows, totalColumns, false, null);

            // Save the cleaned workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Duplicate rows removed successfully. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
