using System;
using Aspose.Cells;

class RemoveCsvDuplicates
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Path for the resulting workbook
        string outputPath = "output.xlsx";

        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV data starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric strings to numbers
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range after import
        int startRow = 0;                         // Assuming the first row contains headers
        int startColumn = 0;
        int endRow = cells.MaxDataRow;            // Last row with data
        int endColumn = cells.MaxDataColumn;      // Last column with data

        // Index of the key column used to identify duplicates (0‑based)
        // Change this value to the appropriate column index for your data
        int keyColumnIndex = 0;

        // Remove duplicate rows based on the key column.
        // The 'hasHeaders' flag is true because the first row is a header.
        // The columnOffsets array specifies which columns participate in duplicate detection.
        cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { keyColumnIndex });

        // Save the cleaned workbook to an XLSX file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}