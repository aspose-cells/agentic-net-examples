using System;
using System.Data;
using Aspose.Cells;

class BulkDataExtractor
{
    // Validates encryption status and extracts data from all worksheets
    public static void ExtractData(string filePath)
    {
        // Detect the file format and check if it is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        if (formatInfo.IsEncrypted)
        {
            Console.WriteLine($"The workbook \"{filePath}\" is encrypted. Extraction aborted.");
            return;
        }

        // Load the workbook (no password required because it is not encrypted)
        Workbook workbook = new Workbook(filePath);

        // Iterate through each worksheet and export its data
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"--- Sheet: {sheet.Name} ---");

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Export the used range to a DataTable (including column names)
            DataTable table = sheet.Cells.ExportDataTable(0, 0, maxRow + 1, maxCol + 1, true);

            // Print the extracted rows
            foreach (DataRow row in table.Rows)
            {
                foreach (object item in row.ItemArray)
                {
                    Console.Write($"{item}\t");
                }
                Console.WriteLine();
            }
        }

        // Clean up resources
        workbook.Dispose();
    }

    static void Main()
    {
        // Example usage
        string path = "input.xlsx";
        ExtractData(path);
    }
}